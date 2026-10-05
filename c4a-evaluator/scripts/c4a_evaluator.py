#!/usr/bin/env python3
# =============================================================================
# c4a_evaluator.py — C4 技能提交自动评审器（Level 1–4 完整流水线）
#
# Author : lsa
# License: MIT
# Python : 3.8+ （核心仅标准库；Excel 输出需 openpyxl，可选）
#
# 用法：
#   python c4a_evaluator.py <FOLDER> [-o OUTDIR] [--format md,json,excel,html]
#                                   [--llm none|auto] [--emit-fixtures]
#
# 设计原则（详见 方案设计.md）：
#   1. 单一事实来源 —— 所有阈值/权重/信号来自 references/rubric.yaml，代码零硬编码
#   2. 规则快筛 + LLM 深审 —— 默认纯规则（确定性、零成本、离线可跑），LLM 为可插拔可选层
#   3. 每条判定可追溯 —— 输出 evidence（命中词/命中文件/未命中项），而非只给分
# =============================================================================
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# 轻量 YAML 读取：优先 pyyaml；无则用内置降级解析器（保证零依赖也能跑）
# ---------------------------------------------------------------------------
try:
    import yaml  # type: ignore
    _HAS_YAML = True
except ImportError:  # pragma: no cover
    _HAS_YAML = False


def _mini_yaml(text: str):
    """极简 YAML 子集解析器：支持嵌套映射、列表(- item)、标量、行内列表 [a, b]。
    仅用于本 rubric 文件结构；保证在无 pyyaml 环境下技能仍可运行。"""
    def coerce(v):
        v = v.strip()
        if v == "" or v.lower() in ("null", "~", "none"):
            return None
        if v.lower() in ("true", "false"):
            return v.lower() == "true"
        if v.startswith("[") and v.endswith("]"):
            inner = v[1:-1].strip()
            if not inner:
                return []
            return [coerce(x) for x in inner.split(",")]
        if (v.startswith('"') and v.endswith('"')) or (v.startswith("'") and v.endswith("'")):
            return v[1:-1]
        try:
            return int(v)
        except ValueError:
            pass
        try:
            return float(v)
        except ValueError:
            pass
        return v

    lines = [l for l in text.splitlines() if l.strip() and not l.lstrip().startswith("#")]
    idx = {"i": 0}

    def block_kind(indent):
        """窥探当前缩进层级是列表还是映射。"""
        j = idx["i"]
        while j < len(lines):
            cur = len(lines[j]) - len(lines[j].lstrip())
            if cur < indent:
                return None
            if cur == indent:
                return "list" if lines[j].strip().startswith("- ") else "map"
            j += 1
        return None

    def parse_block(indent):
        """解析 indent 层级；返回 list 或 dict。列表项若形如 `- k: v`，
        则把该项在更深缩进下的所有 `k: v` 行归入同一个 dict。"""
        kind = block_kind(indent)
        container = [] if kind == "list" else {}
        while idx["i"] < len(lines):
            raw = lines[idx["i"]]
            cur = len(raw) - len(raw.lstrip())
            if cur < indent:
                break
            if cur > indent:
                idx["i"] += 1
                continue
            line = raw.strip()
            if line.startswith("- "):
                if not isinstance(container, list):
                    break
                item = line[2:].strip()
                if ":" in item and not item.startswith(("[", "{", '"', "'")):
                    # 列表项是一个映射：解析该项及其后续同层级的键
                    d = {}
                    item_indent = cur + 2
                    k, _, v = item.partition(":")
                    idx["i"] += 1
                    if v.strip():
                        d[k.strip()] = coerce(v)
                    else:
                        nxt = lines[idx["i"]] if idx["i"] < len(lines) else None
                        if nxt is not None and (len(nxt) - len(nxt.lstrip())) > item_indent:
                            d[k.strip()] = parse_block(item_indent + 2)
                        else:
                            d[k.strip()] = None
                    # 吸收同一列表项的后续键（缩进 = item_indent）
                    while idx["i"] < len(lines):
                        nxt = lines[idx["i"]]
                        nc = len(nxt) - len(nxt.lstrip())
                        if nc != item_indent or nxt.strip().startswith("- "):
                            break
                        k2, _, v2 = nxt.strip().partition(":")
                        idx["i"] += 1
                        if v2.strip():
                            d[k2.strip()] = coerce(v2)
                        else:
                            nn = lines[idx["i"]] if idx["i"] < len(lines) else None
                            if nn is not None and (len(nn) - len(nn.lstrip())) > item_indent:
                                d[k2.strip()] = parse_block(item_indent + 2)
                            else:
                                d[k2.strip()] = None
                    container.append(d)
                else:
                    container.append(coerce(item))
                    idx["i"] += 1
            else:
                if not isinstance(container, dict):
                    break
                k, _, v = line.partition(":")
                idx["i"] += 1
                if v.strip():
                    container[k.strip()] = coerce(v)
                else:
                    nxt = lines[idx["i"]] if idx["i"] < len(lines) else None
                    if nxt is not None and (len(nxt) - len(nxt.lstrip())) > cur:
                        container[k.strip()] = parse_block(cur + 2)
                    else:
                        container[k.strip()] = None
        return container

    return parse_block(0)


def load_rubric(path: Path) -> dict:
    text = path.read_text(encoding="utf-8")
    if _HAS_YAML:
        return yaml.safe_load(text)
    return _mini_yaml(text)


# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
DEFAULT_RUBRIC = SKILL_DIR / "references" / "rubric.yaml"

C4_MARKER_RE = re.compile(r"[_\-](C4A?|c4a?)[_\-]")          # _C4_ / -C4- / _C4A_
AUTHOR_PREFIX_RE = re.compile(r"^([A-Za-z][A-Za-z0-9\-\.]{1,30})[_\-](?:C4A?|c4a?)[_\-]")
VERSION_RE = re.compile(r"[_\-]v(\d+)", re.IGNORECASE)
AUTHOR_IN_TEXT_RE = re.compile(
    r"(?:作者|Author|姓名|By)[\s:：*]*([A-Za-z][A-Za-z0-9\-\. ]{1,30})", re.IGNORECASE)

BINARY_EXTS_DEFAULT = {".png", ".jpg", ".jpeg", ".gif", ".mp4", ".mov", ".webm",
                       ".zip", ".xlsx", ".xls", ".pptx", ".ico", ".webp"}

# 富文本：虽是「二进制容器」，但可经提取层读到正文，参与内容级检测（不跳过）
RICH_TEXT_EXTS = {".docx", ".pdf"}

LEVEL_ICON = {"pass": "✅", "partial": "⚠️", "fail": "❌", "missing": "❌"}


# ===========================================================================
# 模块 1 —— 文件采集与识别（Level 1）
# ===========================================================================
def read_text_safe(path: Path, max_bytes: int = 2_000_000) -> str:
    """安全读取文本文件（UTF-8 → GBK → latin-1 回退）。"""
    try:
        raw = path.read_bytes()[:max_bytes]
    except OSError:
        return ""
    for enc in ("utf-8", "utf-8-sig", "gbk", "gb18030"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="replace")


# ---------------------------------------------------------------------------
# 富文本（.docx / .pdf）正文提取 —— 惰性可选依赖，缺失则优雅降级
#
# 设计要点（对应挑战 L92 / L106 / L109–110「内容级检测」）：
#   · 优先用 python-docx / pypdf 提取正文；
#   · 若未安装，.docx 用纯标准库（zipfile + XML 剥离）兜底 —— 零依赖也能读；
#   · .pdf 无第三方库时无法可靠解析，返回空串并记录 degraded 标记；
#   · 模块级缓存库可用性，避免重复导入开销。
# ---------------------------------------------------------------------------
_DOC_LIB_CACHE: dict = {}


def _lib_available(mod_name: str) -> bool:
    """探测可选依赖是否可导入（结果缓存）。"""
    if mod_name not in _DOC_LIB_CACHE:
        try:
            __import__(mod_name)
            _DOC_LIB_CACHE[mod_name] = True
        except Exception:
            _DOC_LIB_CACHE[mod_name] = False
    return _DOC_LIB_CACHE[mod_name]


def read_docx_text(path: Path, max_chars: int = 400_000) -> str:
    """提取 .docx 正文。
    路径①: python-docx（若有）；路径②: 标准库 zipfile 直接读 word/document.xml 并剥离标签。"""
    # 路径① 首选 python-docx
    if _lib_available("docx"):
        try:
            import docx  # type: ignore
            d = docx.Document(str(path))
            parts = [p.text for p in d.paragraphs if p.text and p.text.strip()]
            for tbl in d.tables:
                for row in tbl.rows:
                    cells = [c.text.strip() for c in row.cells if c.text and c.text.strip()]
                    if cells:
                        parts.append(" | ".join(cells))
            return "\n".join(parts)[:max_chars]
        except Exception:
            pass  # 落到标准库兜底
    # 路径② 纯标准库兜底：docx 本质是 ZIP，正文在 word/document.xml
    try:
        import zipfile
        with zipfile.ZipFile(str(path)) as z:
            xml = ""
            for cand in ("word/document.xml", "word/document2.xml"):
                if cand in z.namelist():
                    xml = z.read(cand).decode("utf-8", errors="replace")
                    break
            if not xml:
                return ""
        # <w:p> 段落 → 换行；<w:t> 文本节点 → 内容；其余标签剥离
        xml = re.sub(r"</w:p>", "\n", xml)
        xml = re.sub(r"<w:tab[^>]*/>", "\t", xml)
        texts = re.findall(r"<w:t[^>]*>(.*?)</w:t>", xml, flags=re.S)
        joined = re.sub(r"<[^>]+>", "", "\n".join(texts) if texts else re.sub(r"<[^>]+>", "", xml))
        return joined[:max_chars]
    except Exception:
        return ""


def read_pdf_text(path: Path, max_chars: int = 400_000) -> str:
    """提取 .pdf 正文（pypdf，可选依赖；缺失则返回空串并降级）。"""
    if not _lib_available("pypdf"):
        return ""
    try:
        from pypdf import PdfReader  # type: ignore
        reader = PdfReader(str(path))
        out = []
        for page in reader.pages[:50]:  # 首 50 页足够覆盖说明文档
            try:
                out.append(page.extract_text() or "")
            except Exception:
                continue
            if sum(len(x) for x in out) > max_chars:
                break
        return "\n".join(out)[:max_chars]
    except Exception:
        return ""


def read_rich_text(rec: dict) -> tuple:
    """按扩展名分派正文提取 → (text, degraded)。
    degraded=True 表示「有该格式但未能读到内容」（如缺库的 .pdf），
    供上层在证据中如实标注，而非静默当作空文件。"""
    ext = rec["ext"]
    p = Path(rec["abs_path"])
    if ext == ".docx":
        t = read_docx_text(p)
        return t, (t == "")
    if ext == ".pdf":
        t = read_pdf_text(p)
        return t, (t == "")
    if ext in (".md", ".markdown", ".txt", ".py", ".json", ".yaml", ".yml",
               ".csv", ".html", ".htm", ".skill", ".log", ".rst", ".toml", ".ini", ".cfg"):
        return read_text_safe(p), False
    return "", True


def extract_any_text(rec: dict) -> str:
    """统一入口：文本类直接读，富文本（docx/pdf）走提取层，二进制返回空串。
    替换原先「is_binary 即跳过」的粗粒度判断，使 .docx/.pdf 可参与内容级检测。"""
    if rec["ext"] in RICH_TEXT_EXTS:
        return read_rich_text(rec)[0]
    if rec["is_binary"]:
        return ""
    return read_text_safe(Path(rec["abs_path"]))


def scan_folder(folder: Path, rubric: dict) -> list:
    """递归扫描，返回文件记录列表（含相对路径/扩展名/大小/时间）。"""
    binary_exts = set(e.lower() for e in rubric.get("limits", {}).get("binary_extensions", [])) or BINARY_EXTS_DEFAULT
    records = []
    for f in sorted(folder.rglob("*")):
        if not f.is_file() or f.name.startswith(".") or f.name.startswith("~$"):
            continue
        try:
            st = f.stat()
        except OSError:
            continue
        ext = f.suffix.lower()
        is_rich = ext in RICH_TEXT_EXTS
        records.append({
            "path": str(f.relative_to(folder)).replace("\\", "/"),
            "abs_path": str(f),
            "name": f.name,
            "stem": f.stem,
            "ext": ext,
            "size_kb": round(st.st_size / 1024, 1),
            "modified": datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d"),
            # 富文本(.docx/.pdf)可提取正文，故 is_binary=False，参与内容级检测
            "is_binary": (ext in binary_exts) and not is_rich,
            "is_rich_text": is_rich,
        })
    return records


def is_c4_file(rec: dict) -> bool:
    """判断是否为 C4 提交文件：文件名含 _C4_ 标记，或位于 C4 相关目录下。"""
    if C4_MARKER_RE.search(rec["name"]):
        return True
    p = rec["path"].lower()
    return any(t in p for t in ("/c4_", "/c4/", "c4提交", "c4_提交", "_c4"))


def extract_author(rec: dict, folder: Path, rubric: dict) -> tuple:
    """作者提取回退链 → (author, method)。"""
    # 1. 文件名前缀（_C4_ 之前）
    m = AUTHOR_PREFIX_RE.match(rec["name"])
    if m:
        return m.group(1), "filename_prefix"
    # 2. 子目录名（若子目录不是通用词）
    generic = {"c4", "c4a", "提交", "submissions", "submission", "materials",
               "outputs", "docs", "doc", "misc", "temp", "tmp", "downloads",
               "download", "desktop", "文件", "杂项", "其他", "other", "others"}
    parts = Path(rec["path"]).parts[:-1]
    for seg in reversed(parts):
        seg_clean = seg.strip()
        if seg_clean and seg_clean.lower() not in generic:
            cand = re.match(r"^([A-Za-z][A-Za-z0-9\-\.]{1,30})$", seg_clean)
            if cand:
                return cand.group(1), "subfolder"
    # 3. 文本内首部（作者/Author: xxx）
    if not rec["is_binary"] and rec["size_kb"] < 5000:
        text = read_text_safe(Path(rec["abs_path"]), 4000)
        m2 = AUTHOR_IN_TEXT_RE.search(text)
        if m2:
            return m2.group(1).strip().split()[0], "content_header"
    return "Unknown", "unmatched"


def parse_version(rec: dict) -> int:
    m = VERSION_RE.search(rec["stem"])
    return int(m.group(1)) if m else 1


def build_bundles(records: list, folder: Path, rubric: dict) -> tuple:
    """按作者分组 → {author: {files, methods}}；同时返回非 C4 文件。
    归属规则：含 _C4_ 标记的文件直接归其作者；同一目录树内与 C4 文件同属一个
    技能包（如 slide-forge/ 内的源文件）的附属文件也一并归入该作者，避免把
    技能包内部文件误报为"非 C4 文件"。"""
    bundles = defaultdict(lambda: {"files": [], "methods": set()})
    non_c4 = []

    # 第一遍：找出所有 C4 标记文件及其作者
    c4_authors = {}
    for rec in records:
        if is_c4_file(rec):
            author, method = extract_author(rec, folder, rubric)
            c4_authors[rec["path"]] = (author, method)

    # 收集这些 C4 文件所在的目录前缀（用于吸收技能包内部附属文件）
    owned_dirs = set()
    for p in c4_authors:
        parts = Path(p).parts[:-1]
        for k in range(1, len(parts) + 1):
            owned_dirs.add("/".join(parts[:k]))

    # 【同目录归属】某个目录若只含单一作者的 C4 文件，则该目录内的其他无标记文件
    # （如交付包顶层的 README.md、说明文件）也归该作者，避免误报为"非 C4 文件"。
    dir_author = {}          # 目录前缀（含空串表示根）→ 唯一作者
    dir_author_multi = set() # 出现多作者冲突的目录
    for p, (a, _) in c4_authors.items():
        prefix = "/".join(Path(p).parts[:-1])
        if prefix in dir_author and dir_author[prefix] != a:
            dir_author_multi.add(prefix)
        else:
            dir_author.setdefault(prefix, a)
    for d in dir_author_multi:
        dir_author.pop(d, None)

    # 识别"解包后的技能包目录"：子目录名与某个 .skill/.zip 的 stem 后缀同名
    # 例：lsa_C4_slide-forge.skill → 包名 slide-forge → 吸收 slide-forge/ 内的文件
    pkg_names = {}
    for p, (a, _) in c4_authors.items():
        stem = Path(p).stem
        m = re.search(r"_C4A?[_-](.+?)(?:_v\d+)?$", stem)
        if m:
            pkg_names[m.group(1).lower()] = a

    # 第二遍：分配
    for rec in records:
        if rec["path"] in c4_authors:
            author, method = c4_authors[rec["path"]]
            bundles[author]["files"].append(rec)
            bundles[author]["methods"].add(method)
            continue
        # 附属文件：① 位于某个含 C4 文件的目录内 → 归入该目录所有者；
        #           ② 位于以某个 .skill/.zip 包名命名的目录内 → 归入该包作者
        parts = Path(rec["path"]).parts[:-1]
        # ⓪ 同目录唯一作者：优先，覆盖顶层 README 这类"与 C4 文件同目录、无标记"的文件
        same_dir = "/".join(parts)
        if same_dir in dir_author:
            bundles[dir_author[same_dir]]["files"].append(rec)
            bundles[dir_author[same_dir]]["methods"].add("same_dir")
            continue
        owner = None
        for k in range(len(parts), 0, -1):
            prefix = "/".join(parts[:k])
            if prefix in owned_dirs:
                owner = k
                break
        if owner is not None:
            dir_prefix = "/".join(parts[:owner])
            cands = [a for p, (a, _) in c4_authors.items()
                     if "/".join(Path(p).parts[:-1]) == dir_prefix]
            if not cands:
                cands = [a for p, (a, _) in c4_authors.items()
                         if p.startswith(dir_prefix + "/")]
            if cands:
                bundles[cands[0]]["files"].append(rec)
                bundles[cands[0]]["methods"].add("package_member")
                continue
        # ② 解包技能包目录
        pkg_owner = None
        for seg in parts:
            if seg.lower() in pkg_names:
                pkg_owner = pkg_names[seg.lower()]
                break
        if pkg_owner:
            bundles[pkg_owner]["files"].append(rec)
            bundles[pkg_owner]["methods"].add("package_member")
            continue
        non_c4.append(rec)
    return bundles, non_c4


# ===========================================================================
# 模块 2 —— 完整性检查（Level 2）
# ===========================================================================
def slot_match_score(rec: dict, slot_cfg: dict, rubric: dict) -> tuple:
    """返回文件对某槽位的匹配等级与特异性 → (level, specificity)。
    level：3=文件名模式命中，2=扩展名+内容信号命中，1=仅扩展名，0=不匹配。
    specificity：命中的文件名模式长度（越长越具体，用于打破平局）。
    设计要点：用 (level, specificity) 二元组排序，避免通用模式（如 "skill"）
    抢走本应属于更具体模式（如 "skill说明"）的槽位。"""
    name_lower = rec["name"].lower()

    # 第一层：文件名模式（最可信）—— 记录最长命中模式长度作为特异性
    best_pat = 0
    for pat in slot_cfg.get("filename_patterns", []):
        if pat.lower() in name_lower:
            best_pat = max(best_pat, len(pat))
    if best_pat:
        return (3, best_pat)

    # 扩展名
    exts = [e.lower() for e in slot_cfg.get("extensions", [])]
    if rec["ext"] not in exts:
        return (0, 0)

    if rec["is_binary"]:
        return (2, 0)

    limit = rubric.get("limits", {}).get("skip_content_over_kb", 50000)
    if rec["size_kb"] > limit:
        return (1, 0)

    signals = slot_cfg.get("content_signals", [])
    if not signals:
        return (1, 0)
    text = extract_any_text(rec).lower()
    n = sum(1 for s in signals if s.lower() in text)
    need = max(2, (len(signals) + 1) // 3)
    if n >= need:
        return (2, 0)
    if len(signals) >= 4 and n >= 2:
        return (1, 0)
    # 扩展名命中但内容信号不足 → 弱匹配(level 1)：作为兜底，供完整性"部分命中"统计，
    # 但会被任何 level≥2 的竞争者优先抢占（避免误判为齐全）。
    return (1, 0) if len(signals) >= 4 else (0, 0)


def match_slot(rec: dict, slot_cfg: dict, rubric: dict) -> bool:
    """布尔判定（供外部/兼容使用）。"""
    return slot_match_score(rec, slot_cfg, rubric)[0] > 0


def check_completeness(files: list, rubric: dict) -> dict:
    """对一位作者的文件做 5 件必交检查。
    采用**贪心一对一分配 + 特异性优先**：按 (匹配等级, 特异性) 降序分配，
    保证①同一文件不被多个槽位抢用；②具体文件名模式（skill说明）优先于通用模式（skill）。"""
    comp_cfg = rubric["completeness"]
    candidates = []
    for slot, cfg in comp_cfg.items():
        for i, rec in enumerate(files):
            lvl, spec = slot_match_score(rec, cfg, rubric)
            if lvl > 0:
                candidates.append((lvl, spec, parse_version(rec), slot, i))
    # 等级 > 特异性 > 版本号(取新) > 槽位名，保证同一槽位优先认领最新版本的文件
    candidates.sort(key=lambda x: (-x[0], -x[1], -x[2], x[3]))

    assigned = {}
    used_files = set()
    for lvl, spec, ver, slot, i in candidates:
        if slot in assigned or i in used_files:
            continue
        assigned[slot] = i
        used_files.add(i)

    result = {}
    for slot, cfg in comp_cfg.items():
        i = assigned.get(slot)
        lvl = slot_match_score(files[i], cfg, rubric) if i is not None else (0, 0)
        result[slot] = {
            "label": cfg["label_cn"],
            "present": i is not None,
            "matched_file": files[i]["path"] if i is not None else None,
            "confidence": lvl[0],
        }
    n_present = sum(1 for v in result.values() if v["present"])
    total = len(comp_cfg)
    if n_present == total:
        status = "complete"          # ✅ 齐全
    elif n_present >= 3:
        status = "partial"           # ⚠️ 部分缺失
    else:
        status = "insufficient"      # ❌ 严重缺失
    return {
        "slots": result,
        "present": n_present,
        "total": total,
        "status": status,
        "score": round(n_present / total, 4),
        "missing": [v["label"] for v in result.values() if not v["present"]],
        "low_confidence": [v["label"] for v in result.values()
                           if v["present"] and v["confidence"] < 2],
    }


# ===========================================================================
# 模块 3 —— 质量评审（Level 3：规则引擎 + 可插拔 LLM）
# ===========================================================================
def gather_text(files: list, filled_files: set) -> str:
    """汇总该作者提交包的全文（文本类 + 富文本 docx/pdf），供质量评审使用。"""
    chunks = []
    for rec in files:
        # 跳过纯二进制（图片/视频/压缩包）；.docx/.pdf 已不属于 is_binary，会走提取层
        if rec["is_binary"] or rec["size_kb"] > 50000:
            continue
        text = extract_any_text(rec)
        if not text:
            continue
        chunks.append(f"\n\n===== FILE: {rec['path']} =====\n{text}")
    return "".join(chunks)


def python_syntax_ok(rec: dict) -> bool:
    """对 .py 文件做静态语法检查。"""
    if rec["ext"] != ".py":
        return True
    try:
        import ast
        src = read_text_safe(Path(rec["abs_path"]))
        ast.parse(src)
        return True
    except SyntaxError:
        return False
    except Exception:
        return True


def count_hits(text_lower: str, signals: list) -> tuple:
    """返回 (命中数, 命中词列表)。"""
    hits = []
    for s in signals:
        if s.lower() in text_lower:
            hits.append(s)
    return len(hits), hits


def eval_rule(rule: dict, text: str, ctx: dict) -> dict:
    """执行单条规则，返回 {score(0-1), evidence, detail}。"""
    kind = rule["kind"]
    text_lower = text.lower()
    rid = rule["id"]

    # 条件性规则：仅当包内有对应扩展名时才评估
    only_if = rule.get("only_if_present")
    if only_if:
        if not any(ctx["exts"] & {e.lower() for e in only_if}):
            return {"score": None, "skipped": True,
                    "evidence": f"跳过（包内无 {'/'.join(only_if)} 文件）"}

    if kind == "must_have":
        n, hits = count_hits(text_lower, rule["signals"])
        min_hits = rule.get("min_hits", 1)
        if n >= min_hits:
            score = 1.0
        elif n >= 1:
            score = rule.get("key_partial", 0.5)
        else:
            score = 0.0
        ev = f"命中 {n}/{min_hits} 个信号：{', '.join(hits[:6])}" if hits else "未命中任何信号"
        return {"score": score, "evidence": ev, "hits": hits}

    if kind == "must_avoid":
        n, hits = count_hits(text_lower, rule["signals"])
        score = max(0.0, 1.0 - n * rule.get("penalize_each", 0.34))
        ev = f"检出 {n} 处负向信号" + (f"：{', '.join(hits[:6])}" if hits else "（无，合规）")
        return {"score": score, "evidence": ev, "hits": hits}

    if kind == "pattern":
        pat = rule["pattern"]
        found = re.findall(pat, text, re.IGNORECASE)
        score = 1.0 if len(found) >= rule.get("min_hits", 1) else 0.0
        return {"score": score,
                "evidence": f"正则命中 {len(found)} 处" + ("，例：'输入…输出…'" if found else "（未命中）"),
                "hits": found}

    if kind == "count":
        n, hits = count_hits(text_lower, rule["signals"])
        min_hits = rule.get("min_hits", 1)
        score = min(1.0, n / min_hits) if min_hits else 0.0
        return {"score": score, "evidence": f"计数 {n}/{min_hits}", "hits": hits}

    if kind == "ai_assisted":
        # 规则无法可靠判定 → 交 LLM / 人工；默认宽容分并标记
        return {"score": rule.get("leniency", 0.6), "ai_assisted": True,
                "evidence": "需 LLM/人工语义复核（规则不判定）"}

    return {"score": 0.0, "evidence": f"未知规则类型 {kind}"}


def eval_criterion(crit_id: str, cfg: dict, text: str, ctx: dict) -> dict:
    """评估单个质量维度，返回 score/rating/evidence/rules。"""
    total_w, acc_w = 0.0, 0.0
    rule_reports = []
    for rule in cfg["rules"]:
        res = eval_rule(rule, text, ctx)
        rule_reports.append({
            "id": rule["id"],
            "desc": rule["description"],
            "kind": rule["kind"],
            "weight": rule["weight"],
            "score": res["score"],
            "evidence": res["evidence"],
            "skipped": res.get("skipped", False),
        })
        if res.get("skipped"):
            continue
        total_w += rule["weight"]
        acc_w += rule["weight"] * res["score"]

    raw = (acc_w / total_w) if total_w else 0.0
    if raw >= cfg["pass_threshold"]:
        rating = "pass"
    elif raw >= cfg["partial_threshold"]:
        rating = "partial"
    else:
        rating = "fail"
    return {
        "id": crit_id,
        "label": cfg["label_cn"],
        "label_en": cfg["label_en"],
        "raw": round(raw, 4),
        "rating": rating,
        "icon": LEVEL_ICON[rating],
        "weight": cfg["weight"],
        "rules": rule_reports,
    }


def evaluate_quality(files: list, comp: dict, rubric: dict, llm=None) -> dict:
    """C4 四条件质量评审。"""
    filled = {v["matched_file"] for v in comp["slots"].values() if v["matched_file"]}
    text = gather_text(files, filled)

    # 注入代码语法检查结果（负向信号）
    syntax_bad = [r["path"] for r in files if r["ext"] == ".py" and not python_syntax_ok(r)]
    if syntax_bad:
        text += "\n__SYNTAX_ERROR__ " + " ".join(syntax_bad)

    ctx = {"exts": {r["ext"] for r in files}, "n_files": len(files),
           "syntax_bad": syntax_bad}

    crits = {}
    for cid, cfg in rubric["quality_criteria"].items():
        crits[cid] = eval_criterion(cid, cfg, text, ctx)

    # 可选 LLM 深审层（可插拔；默认关闭）
    llm_notes = []
    if llm is not None:
        try:
            llm_notes = llm.review(text, crits)
        except Exception as e:  # noqa: BLE001
            llm_notes = [f"[LLM 层异常，已回退纯规则] {type(e).__name__}: {e}"]

    # 质量分（0–4，每维度 ✅=1 ⚠️=0.5 ❌=0）
    val = {"pass": 1.0, "partial": 0.5, "fail": 0.0}
    q_scores = {cid: val[c["rating"]] for cid, c in crits.items()}
    # quality_score 采用 0–4 分制（四条件各 0 / 0.5 / 1），与报告中的 "/4" 口径一致
    quality_score = round(sum(q_scores.values()), 4) if q_scores else 0.0
    return {"criteria": crits, "quality_score": quality_score, "llm_notes": llm_notes,
            "syntax_errors": syntax_bad}


# ---------------------------------------------------------------------------
# 可插拔 LLM 层（示例骨架；真实环境注入实现，如调用本机 Agent）
# ---------------------------------------------------------------------------
class LLMReviewer:  # pragma: no cover
    """LLM 深审层接口。子类实现 review(text, crits) -> List[str]。
    设计为可选：未注入则技能以纯规则运行（确定性、离线、零成本）。"""

    def review(self, text: str, crits: dict):
        raise NotImplementedError


# ===========================================================================
# 模块 4 —— 综合评分 / 排名 / 版本追踪（Level 4）
# ===========================================================================
def composite_score(comp: dict, quality: dict, rubric: dict) -> float:
    cc = rubric["composite"]
    cw = cc["completeness_weight"]
    qw = cc["quality_weight"]
    q = quality["quality_score"] / 4.0
    return round(comp["score"] * cw + q * qw, 4)


def make_suggestions(comp: dict, quality: dict, author: str) -> list:
    """针对每位作者生成个性化『下一步行动』。"""
    sug = []
    if comp["missing"]:
        sug.append(f"**补齐缺失文件**：{'、'.join(comp['missing'])} —— 这是完整性硬门槛，直接决定能否进入质量评审。")
    c = quality["criteria"]
    weak = [v for v in c.values() if v["rating"] != "pass"]
    for w in weak:
        worst = None
        for r in w["rules"]:
            if r.get("skipped"):
                continue
            if worst is None or (r["score"] or 0) < (worst["score"] or 0):
                worst = r
        if worst and worst["score"] is not None and worst["score"] < 1.0:
            sug.append(f"**加强「{w['label']}」**（当前 {w['icon']}）：最弱子项是「{worst['desc']}」——{worst['evidence']}。")
    if quality["syntax_errors"]:
        sug.append(f"**修复语法错误**：{', '.join(quality['syntax_errors'])} 无法通过 py_compile，评审器已扣分。")
    if not sug:
        sug.append("四个条件全部 ✅，建议补充更多真实使用案例、提升被使用次数（C4 第一评判标准）。")
    return sug


def evaluate_author(author: str, bundle: dict, rubric: dict, llm=None) -> dict:
    files = bundle["files"]
    comp = check_completeness(files, rubric)
    quality = evaluate_quality(files, comp, rubric, llm=llm)
    versions = sorted({parse_version(r) for r in files})
    return {
        "author": author,
        "n_files": len(files),
        "files": [{"path": r["path"], "ext": r["ext"], "size_kb": r["size_kb"],
                   "modified": r["modified"], "version": parse_version(r)} for r in files],
        "author_method": sorted(bundle["methods"]),
        "completeness": comp,
        "quality": quality,
        "composite": composite_score(comp, quality, rubric),
        "versions": versions,
        "latest_version": max(versions) if versions else 1,
        "suggestions": make_suggestions(comp, quality, author),
    }


def build_version_trail(results: list) -> dict:
    """展示每位作者的迭代轨迹（v1 → v2 → ...）。"""
    trail = {}
    for r in results:
        if len(r["versions"]) > 1:
            trail[r["author"]] = r["versions"]
    return trail


# ===========================================================================
# 报告生成（Markdown / JSON / Excel / HTML）
# ===========================================================================
STATUS_ICON = {"complete": "✅ 齐全", "partial": "⚠️ 部分缺失", "insufficient": "❌ 严重缺失"}
CRIT_ORDER = ["reusable", "executable", "verifiable", "clear_io"]


def gen_markdown(ctx: dict) -> str:
    results = ctx["results"]
    n = len(results)
    complete = sum(1 for r in results if r["completeness"]["status"] == "complete")
    partial = sum(1 for r in results if r["completeness"]["status"] == "partial")
    insuff = n - complete - partial
    avg_q = round(sum(r["quality"]["quality_score"] for r in results) / n, 2) if n else 0
    avg_c = round(sum(r["composite"] for r in results) / n, 4) if n else 0

    L = []
    L.append("# C4 提交自动评审报告\n")
    L.append(f"> 生成工具：**c4a-evaluator v{ctx['meta']['version']}** ｜ 评审模式：**{ctx['meta']['mode']}**")
    L.append(f"> 生成时间：{ctx['meta']['generated_at']}")
    L.append(f"> 扫描路径：`{ctx['meta']['folder']}` ｜ 识别提交：**{n}** 位作者 / **{ctx['meta']['n_files']}** 个文件\n")
    L.append("---\n")

    # 一、班级总览
    L.append("## 一、班级总览（Level 4 仪表板）\n")
    L.append("| 指标 | 数值 |")
    L.append("| --- | --- |")
    L.append(f"| 提交人数 | {n} |")
    L.append(f"| 完整提交（5/5） | {complete} |")
    L.append(f"| 部分提交（3–4/5） | {partial} |")
    L.append(f"| 严重缺失（<3/5） | {insuff} |")
    L.append(f"| 平均质量分 | {avg_q}/4.0 |")
    L.append(f"| 平均综合分 | {avg_c} |\n")

    # 缺失分布
    miss_counter = defaultdict(int)
    for r in results:
        for m in r["completeness"]["missing"]:
            miss_counter[m] += 1
    if miss_counter:
        L.append("**全班最常缺失的文件：**\n")
        L.append("| 缺失文件 | 缺失人数 |")
        L.append("| --- | --- |")
        for k, v in sorted(miss_counter.items(), key=lambda x: -x[1]):
            L.append(f"| {k} | {v} |")
        L.append("")

    # 质量分布
    L.append("**四条件通过率：**\n")
    L.append("| 维度 | ✅ | ⚠️ | ❌ |")
    L.append("| --- | --- | --- | --- |")
    for cid in CRIT_ORDER:
        p = sum(1 for r in results if r["quality"]["criteria"][cid]["rating"] == "pass")
        pa = sum(1 for r in results if r["quality"]["criteria"][cid]["rating"] == "partial")
        fa = n - p - pa
        L.append(f"| {ctx['rubric']['quality_criteria'][cid]['label_cn']} | {p} | {pa} | {fa} |")
    L.append("")

    # 二、排名
    L.append("## 二、综合排名\n")
    L.append(f"> 综合分 = 完整性 × {ctx['rubric']['composite']['completeness_weight']} + 质量分(归一) × {ctx['rubric']['composite']['quality_weight']}\n")
    L.append("| 排名 | 作者 | 完整性 | 质量分 | 综合分 | 四条件 |")
    L.append("| --- | --- | --- | --- | --- | --- |")
    ranked = sorted(results, key=lambda r: -r["composite"])
    for i, r in enumerate(ranked, 1):
        icons = "".join(r["quality"]["criteria"][c]["icon"] for c in CRIT_ORDER)
        L.append(f"| {i} | **{r['author']}** | {r['completeness']['present']}/5 | "
                 f"{r['quality']['quality_score']}/4 | **{r['composite']}** | {icons} |")
    L.append("")

    # 版本追踪
    trail = build_version_trail(results)
    if trail:
        L.append("## 三、版本追踪\n")
        for a, vs in trail.items():
            L.append(f"- **{a}**：{' → '.join('v' + str(v) for v in vs)}")
        L.append("")

    # 四、作者详情
    L.append("## 四、作者详情\n")
    for r in ranked:
        L.append(f"### {r['author']}\n")
        L.append(f"- 提交文件：{r['n_files']} 个 ｜ 作者识别方式：{', '.join(r['author_method'])} ｜ 版本：v{r['latest_version']}\n")
        L.append("**完整性检查：**\n")
        L.append("| 必交文件 | 状态 | 匹配文件 |")
        L.append("| --- | --- | --- |")
        for slot, v in r["completeness"]["slots"].items():
            icon = "✅" if v["present"] else "❌"
            L.append(f"| {v['label']} | {icon} | {v['matched_file'] or '—'} |")
        L.append(f"\n**完整性结论：{STATUS_ICON[r['completeness']['status']]}**"
                 + (f"（缺：{'、'.join(r['completeness']['missing'])}）\n" if r["completeness"]["missing"] else "\n"))
        L.append("**质量评审（C4 四条件）：**\n")
        L.append("| 条件 | 评级 | 分 | 关键依据 |")
        L.append("| --- | --- | --- | --- |")
        for cid in CRIT_ORDER:
            c = r["quality"]["criteria"][cid]
            key_ev = next((rr["evidence"] for rr in c["rules"] if rr.get("score") is not None
                           and rr["score"] < 1.0), c["rules"][0]["evidence"])
            L.append(f"| {c['label']} | {c['icon']} | {c['raw']} | {key_ev} |")
        L.append("")
        # 证据明细
        L.append("<details><summary>展开逐规则证据</summary>\n")
        for cid in CRIT_ORDER:
            c = r["quality"]["criteria"][cid]
            L.append(f"- **{c['label']}**（{c['icon']} raw={c['raw']}）")
            for rr in c["rules"]:
                if rr.get("skipped"):
                    L.append(f"  - `{rr['id']}` — {rr['evidence']}")
                else:
                    L.append(f"  - `{rr['id']}` [{rr['score']}] {rr['desc']} → {rr['evidence']}")
        L.append("\n</details>\n")
        L.append("**改进建议：**\n")
        for i, s in enumerate(r["suggestions"], 1):
            L.append(f"{i}. {s}")
        L.append("\n---\n")

    # 五、全班改进建议
    L.append("## 五、全班改进建议\n")
    if miss_counter:
        top_miss = max(miss_counter.items(), key=lambda x: x[1])
        L.append(f"- **最常见缺失**：{top_miss[0]}（{top_miss[1]} 人缺失）")
    weak_crit = min(CRIT_ORDER,
                    key=lambda cid: sum(1 for r in results if r["quality"]["criteria"][cid]["rating"] == "pass"))
    L.append(f"- **最弱维度**：{ctx['rubric']['quality_criteria'][weak_crit]['label_cn']}")
    L.append("- 建议下次提交前**用本技能自检**：`python c4a_evaluator.py <你的文件夹>`，先过完整性再进质量。")
    L.append("- 团体共性：把「输入/输出」写成一句话模板、补一张 demo 截图、环境要求写明 Python 版本，即可显著提分。")
    L.append("")

    if ctx.get("non_c4"):
        L.append("## 六、非 C4 文件（已排除）\n")
        L.append("| 文件 | 说明 |")
        L.append("| --- | --- |")
        for r in ctx["non_c4"][:30]:
            L.append(f"| `{r['path']}` | 未含 _C4_ 标记，未纳入评审 |")
        L.append("")
    L.append("\n> 本报告由 c4a-evaluator 自动生成；规则评分确定可复现，语义判定建议人工抽检。")
    return "\n".join(L)


def gen_excel(ctx: dict, out_path: Path):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("[warn] openpyxl 未安装，跳过 Excel 输出。", file=sys.stderr)
        return
    wb = Workbook()
    hfont = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    hfill = PatternFill("solid", start_color="2F5496", end_color="2F5496")
    bd = Border(*[Side(style="thin")] * 4)
    cfont = Font(name="Arial", size=10)

    # Sheet1 总览
    ws = wb.active
    ws.title = "班级总览"
    headers = ["排名", "作者", "文件数", "完整性", "可复用", "可执行", "可验证", "IO明确", "质量分", "综合分", "缺失文件"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h); c.font = hfont; c.fill = hfill
        c.alignment = Alignment("center", wrap_text=True); c.border = bd
    for i, r in enumerate(sorted(ctx["results"], key=lambda x: -x["composite"]), 2):
        vals = [i - 1, r["author"], r["n_files"],
                f"{r['completeness']['present']}/5"]
        vals += [r["quality"]["criteria"][c]["icon"] for c in CRIT_ORDER]
        vals += [r["quality"]["quality_score"], r["composite"],
                 "、".join(r["completeness"]["missing"]) or "—"]
        for j, v in enumerate(vals, 1):
            c = ws.cell(i, j, v); c.font = cfont; c.border = bd
    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16
    ws.freeze_panes = "A2"

    # Sheet2 明细
    ws2 = wb.create_sheet("评审明细")
    h2 = ["作者", "维度", "评级", "归一分明细"]
    for i, h in enumerate(h2, 1):
        c = ws2.cell(1, i, h); c.font = hfont; c.fill = hfill; c.border = bd
    row = 2
    for r in ctx["results"]:
        for cid in CRIT_ORDER:
            c = r["quality"]["criteria"][cid]
            detail = " ; ".join(f"{rr['id']}={rr['score']}" for rr in c["rules"] if not rr.get("skipped"))
            for j, v in enumerate([r["author"], c["label"], c["icon"], detail], 1):
                cc = ws2.cell(row, j, v); cc.font = cfont; cc.border = bd
            row += 1
    for col in range(1, len(h2) + 1):
        ws2.column_dimensions[get_column_letter(col)].width = 40

    # Sheet3 文件清单
    ws3 = wb.create_sheet("文件清单")
    h3 = ["作者", "文件", "扩展名", "大小KB", "修改日期", "版本"]
    for i, h in enumerate(h3, 1):
        c = ws3.cell(1, i, h); c.font = hfont; c.fill = hfill; c.border = bd
    row = 2
    for r in ctx["results"]:
        for f in r["files"]:
            for j, v in enumerate([r["author"], f["path"], f["ext"], f["size_kb"], f["modified"], f["version"]], 1):
                cc = ws3.cell(row, j, v); cc.font = cfont; cc.border = bd
            row += 1
    for col in range(1, len(h3) + 1):
        ws3.column_dimensions[get_column_letter(col)].width = 30

    wb.save(str(out_path))
    print(f"[ok] Excel → {out_path}")


def gen_html(ctx: dict) -> str:
    results = sorted(ctx["results"], key=lambda r: -r["composite"])
    n = len(results)
    avg_q = round(sum(r["quality"]["quality_score"] for r in results) / n, 2) if n else 0
    complete = sum(1 for r in results if r["completeness"]["status"] == "complete")

    rows = []
    for i, r in enumerate(results, 1):
        icons = "".join(f"<span class='ic'>{r['quality']['criteria'][c]['icon']}</span>" for c in CRIT_ORDER)
        rows.append(f"""<tr>
          <td class="rk">{i}</td><td class="au">{r['author']}</td>
          <td>{r['completeness']['present']}/5</td>
          <td>{r['quality']['quality_score']}/4</td>
          <td class="cs">{r['composite']}</td><td class="ics">{icons}</td>
          <td class="ms">{'、'.join(r['completeness']['missing']) or '—'}</td></tr>""")
    bars = []
    for cid in CRIT_ORDER:
        p = sum(1 for r in results if r["quality"]["criteria"][cid]["rating"] == "pass")
        pct = round(100 * p / n) if n else 0
        bars.append(f"""<div class="bar"><div class="bl">{ctx['rubric']['quality_criteria'][cid]['label_cn']}</div>
          <div class="track"><div class="fill" style="width:{pct}%"></div></div><div class="pv">{pct}%</div></div>""")

    return f"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<title>C4 提交自动评审 · 班级仪表板</title>
<style>
:root{{--bg:#f5f7fb;--card:#fff;--ink:#1a2233;--mut:#6b7688;--line:#e3e8f0;--pri:#3b5bdb;--ok:#2f9e44;--warn:#e8a020;--bad:#e03131}}
*{{box-sizing:border-box}}body{{margin:0;font-family:"Microsoft YaHei","PingFang SC",system-ui,sans-serif;background:var(--bg);color:var(--ink);padding:36px}}
.wrap{{max-width:1080px;margin:0 auto}}
h1{{font-size:26px;margin:0 0 6px}}.sub{{color:var(--mut);font-size:13px;margin-bottom:24px}}
.kpis{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:24px}}
.kpi{{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px}}
.kpi .v{{font-size:26px;font-weight:700;color:var(--pri)}}.kpi .l{{font-size:12px;color:var(--mut);margin-top:4px}}
.card{{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:20px;margin-bottom:20px}}
h2{{font-size:16px;margin:0 0 14px}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
th,td{{padding:9px 10px;border-bottom:1px solid var(--line);text-align:left}}
th{{color:var(--mut);font-weight:600;font-size:12px;background:#fafbfe}}
.rk{{color:var(--mut)}}.au{{font-weight:700}}.cs{{font-weight:700;color:var(--pri)}}
.ic{{font-size:14px}}.ms{{color:var(--bad);font-size:12px}}
.bar{{display:flex;align-items:center;gap:12px;margin:8px 0}}
.bl{{width:90px;font-size:13px;color:var(--mut)}}.track{{flex:1;height:10px;background:#eef1f7;border-radius:6px;overflow:hidden}}
.fill{{height:100%;background:linear-gradient(90deg,#4c6ef5,#748ffc)}}.pv{{width:44px;text-align:right;font-size:13px;font-weight:600}}
</style></head><body><div class="wrap">
<h1>C4 提交自动评审 · 班级仪表板</h1>
<div class="sub">c4a-evaluator v{ctx['meta']['version']} ｜ {ctx['meta']['mode']} ｜ 生成于 {ctx['meta']['generated_at']} ｜ 扫描 {ctx['meta']['n_files']} 个文件</div>
<div class="kpis">
  <div class="kpi"><div class="v">{n}</div><div class="l">提交人数</div></div>
  <div class="kpi"><div class="v">{complete}</div><div class="l">完整提交 5/5</div></div>
  <div class="kpi"><div class="v">{avg_q}<span style="font-size:15px;color:var(--mut)">/4</span></div><div class="l">平均质量分</div></div>
  <div class="kpi"><div class="v">{round(100*complete/n) if n else 0}%</div><div class="l">完整率</div></div>
</div>
<div class="card"><h2>四条件通过率</h2>{''.join(bars)}</div>
<div class="card"><h2>综合排名</h2>
<table><thead><tr><th>#</th><th>作者</th><th>完整性</th><th>质量分</th><th>综合分</th><th>可复用/可执行/可验证/IO</th><th>缺失</th></tr></thead>
<tbody>{''.join(rows)}</tbody></table></div>
<div class="card"><h2>说明</h2><p style="font-size:13px;color:var(--mut);line-height:1.7">
综合分 = 完整性 × {ctx['rubric']['composite']['completeness_weight']} + 质量分(归一) × {ctx['rubric']['composite']['quality_weight']}。
评级：✅ 达标 / ⚠️ 部分 / ❌ 缺失。全部判定均附证据，规则口径见 references/rubric.yaml。
本页面为可选可视化输出，与 Markdown / Excel 报告同源。</p></div>
</div></body></html>"""


# ===========================================================================
# 主流程
# ===========================================================================
def run(folder: Path, rubric: dict, out_dir: Path, formats: list, llm=None) -> dict:
    records = scan_folder(folder, rubric)
    bundles, non_c4 = build_bundles(records, folder, rubric)
    results = [evaluate_author(a, b, rubric, llm=llm) for a, b in sorted(bundles.items())]
    ctx = {
        "meta": {
            "version": rubric["meta"]["version"],
            "mode": "规则 + LLM 深审" if llm else "纯规则（确定性）",
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "folder": str(folder),
            "n_files": len(records),
        },
        "rubric": rubric,
        "results": results,
        "non_c4": [{"path": r["path"]} for r in non_c4],
    }
    out_dir.mkdir(parents=True, exist_ok=True)

    if "json" in formats:
        (out_dir / "c4a_eval_result.json").write_text(
            json.dumps({k: v for k, v in ctx.items() if k != "rubric"}, ensure_ascii=False, indent=2),
            encoding="utf-8")
        print(f"[ok] JSON  → {out_dir / 'c4a_eval_result.json'}")
    if "md" in formats:
        (out_dir / "c4a_eval_report.md").write_text(gen_markdown(ctx), encoding="utf-8")
        print(f"[ok] MD    → {out_dir / 'c4a_eval_report.md'}")
    if "excel" in formats:
        gen_excel(ctx, out_dir / "c4a_eval_report.xlsx")
    if "html" in formats:
        (out_dir / "c4a_eval_dashboard.html").write_text(gen_html(ctx), encoding="utf-8")
        print(f"[ok] HTML  → {out_dir / 'c4a_eval_dashboard.html'}")
    return ctx


def main():
    ap = argparse.ArgumentParser(description="c4a-evaluator — C4 技能提交自动评审器")
    ap.add_argument("folder", help="待评审的提交文件夹路径（WeChat 同步文件夹 / 下载文件夹均可）")
    ap.add_argument("-o", "--outdir", default="./c4a_eval_out", help="输出目录")
    ap.add_argument("--format", default="md,json", help="输出格式，逗号分隔：md,json,excel,html")
    ap.add_argument("--rubric", default=str(DEFAULT_RUBRIC), help="评审口径 YAML 路径")
    ap.add_argument("--llm", default="none", choices=["none", "auto"],
                    help="LLM 深审层：none=纯规则（默认）")
    args = ap.parse_args()

    folder = Path(args.folder).expanduser().resolve()
    if not folder.is_dir():
        print(f"ERROR: '{folder}' 不是目录。", file=sys.stderr)
        sys.exit(1)

    rubric = load_rubric(Path(args.rubric))
    formats = [f.strip() for f in args.format.split(",") if f.strip()]
    ctx = run(folder, rubric, Path(args.outdir), formats, llm=None)

    n = len(ctx["results"])
    print(f"\n=== 评审完成：识别 {n} 位作者 / {ctx['meta']['n_files']} 个文件 ===")
    for r in sorted(ctx["results"], key=lambda x: -x["composite"]):
        icons = "".join(r["quality"]["criteria"][c]["icon"] for c in CRIT_ORDER)
        print(f"  {r['author']:<16} 完整性 {r['completeness']['present']}/5  "
              f"质量 {r['quality']['quality_score']}/4  综合 {r['composite']}  {icons}")


if __name__ == "__main__":
    main()
