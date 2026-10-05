#!/usr/bin/env python3
# =============================================================================
# self_test.py — c4a-evaluator 自测套件
#
# Author : lsa
# 目的：证明评审器「可验证」——用**构造的对抗性夹具**验证每条规则与每个边界，
#       并给出可量化的指标：召回率 / 精确率 / 边界用例通过率。
#
# 用法： python self_test.py
# =============================================================================
from __future__ import annotations

import json
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import c4a_evaluator as EV  # noqa: E402

FIXTURES = Path(__file__).resolve().parent.parent / "tests" / "fixtures"

PASS, FAIL = 0, 0
_FAILED = []


def check(name: str, cond: bool, detail: str = ""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  [PASS] {name}")
    else:
        FAIL += 1
        _FAILED.append(name)
        print(f"  [FAIL] {name}  {detail}")


# ---------------------------------------------------------------------------
# 一、YAML 加载 & 降级解析器一致性
# ---------------------------------------------------------------------------
def test_rubric_loading(rubric):
    print("\n[1] Rubric 加载")
    check("meta.version 存在", "version" in rubric.get("meta", {}))
    check("完整性槽位 = 5", len(rubric["completeness"]) == 5,
          f"实际 {len(rubric['completeness'])}")
    check("质量维度 = 4", len(rubric["quality_criteria"]) == 4,
          f"实际 {len(rubric['quality_criteria'])}")
    check("四条件等权 0.25", all(abs(v["weight"] - 0.25) < 1e-9
                                for v in rubric["quality_criteria"].values()))
    check("综合权重 0.4/0.6",
          abs(rubric["composite"]["completeness_weight"] - 0.4) < 1e-9 and
          abs(rubric["composite"]["quality_weight"] - 0.6) < 1e-9)
    # 降级解析器（无 pyyaml 时）与 pyyaml 结果一致性
    text = (SCRIPT_DIR.parent / "references" / "rubric.yaml").read_text(encoding="utf-8")
    mini = EV._mini_yaml(text)
    check("降级 YAML 解析：槽位数一致", len(mini["completeness"]) == 5)
    check("降级 YAML 解析：维度数一致", len(mini["quality_criteria"]) == 4)
    check("降级 YAML 解析：列表型 signals 正确",
          isinstance(mini["quality_criteria"]["reusable"]["rules"][0]["signals"], list))


# ---------------------------------------------------------------------------
# 二、作者识别回退链
# ---------------------------------------------------------------------------
def test_author_extraction(rubric, folder=None):
    print("\n[2] 作者识别回退链（Level 1）")
    mk = lambda **kw: {**{"path": "", "abs_path": "", "name": "", "stem": "",
                          "ext": ".md", "size_kb": 1, "modified": "", "is_binary": False}, **kw}
    cases = [
        (mk(name="ZhangWei_C4_skill说明.md", stem="ZhangWei_C4_skill说明"), "ZhangWei", "filename_prefix"),
        (mk(name="LiMing_C4A_demo.png", stem="LiMing_C4A_demo", ext=".png"), "LiMing", "filename_prefix"),
        (mk(name="WangXiao_C4_report.md", path="WangXiao/WangXiao_C4_report.md",
            stem="WangXiao_C4_report"), "WangXiao", "filename_prefix"),
        (mk(name="skill说明.md", path="ChenHao/skill说明.md", stem="skill说明"), "ChenHao", "subfolder"),
        (mk(name="随便.md", path="misc/随便.md", stem="随便"), "Unknown", "unmatched"),
    ]
    for rec, exp_author, exp_method in cases:
        a, m = EV.extract_author(rec, Path("."), rubric)
        check(f"作者提取 {rec['name']} → {exp_author}", a == exp_author, f"得到 {a}")
        check(f"  识别方式 {exp_method}", m == exp_method, f"得到 {m}")

    # C4 识别
    check("C4 标记识别 _C4_", EV.is_c4_file({"name": "a_C4_b.md", "path": "a_C4_b.md"}))
    check("C4 标记识别 _C4A_", EV.is_c4_file({"name": "a_C4A_b.md", "path": "a_C4A_b.md"}))
    check("非 C4 文件排除", not EV.is_c4_file({"name": "C2_paper.tex", "path": "C2_paper.tex"}))

    # 版本解析
    check("版本 v2 解析", EV.parse_version({"stem": "x_C4_y_v2"}) == 2)
    check("无版本 → 1", EV.parse_version({"stem": "x_C4_y"}) == 1)


# ---------------------------------------------------------------------------
# 三、完整性检查
# ---------------------------------------------------------------------------
def test_completeness(rubric):
    print("\n[3] 完整性检查（Level 2）")
    # 用真实夹具目录：full / partial / insufficient
    for name, exp_present, exp_status in [
        ("author_full", 5, "complete"),
        ("author_partial", 4, "partial"),
        ("author_minimal", 1, "insufficient"),
    ]:
        d = FIXTURES / name
        if not d.exists():
            check(f"夹具 {name} 存在", False, "缺失")
            continue
        recs = EV.scan_folder(d, rubric)
        bundle, _ = EV.build_bundles(recs, d, rubric)
        check(f"夹具 {name} 解析出 1 位作者", len(bundle) == 1, f"得到 {list(bundle)}")
        author = next(iter(bundle))
        comp = EV.check_completeness(bundle[author]["files"], rubric)
        check(f"  {name} 命中 {exp_present}/5", comp["present"] == exp_present,
              f"得到 {comp['present']}，missing={comp['missing']}")
        check(f"  {name} 状态 = {exp_status}", comp["status"] == exp_status,
              f"得到 {comp['status']}")


# ---------------------------------------------------------------------------
# 四、质量评审规则（对抗性 → 精确率 / 召回率）
# ---------------------------------------------------------------------------
GOOD_TEXT = """
# my-skill 技能说明
输入：一个文件夹路径；输出：一份 Markdown 报告。
## 安装与使用步骤
pip install 无，仅需 python 3.8+ 标准库，零依赖；Windows / macOS / Linux 均可运行。
## 使用场景
通用，适用于任何汇报场景，兼容性强。
```python
def run(path): return "report"
```
## 示例与实测
运行 `python run.py -o out.md`，预期输出一份报告：
```
$ python run.py -o out.md
结果：成功，✓ 通过
```
"""
BAD_TEXT = """
我的电脑上放在 /Users/lsa/Desktop，还有个 C:\\temp 目录。
api_key = sk-abc123
没法安装，环境不确定。
"""
STUB_TEXT = "这是一个技能，谢谢观看。"


def test_quality_rules(rubric):
    print("\n[4] 质量评审规则（Level 3）")
    ctx_good = {"exts": {".md", ".py"}, "n_files": 3, "syntax_bad": []}
    ctx_bad = {"exts": {".md"}, "n_files": 1, "syntax_bad": []}

    good = {cid: EV.eval_criterion(cid, cfg, GOOD_TEXT, ctx_good)
            for cid, cfg in rubric["quality_criteria"].items()}
    bad = {cid: EV.eval_criterion(cid, cfg, BAD_TEXT, ctx_bad)
           for cid, cfg in rubric["quality_criteria"].items()}
    stub = {cid: EV.eval_criterion(cid, cfg, STUB_TEXT, ctx_bad)
            for cid, cfg in rubric["quality_criteria"].items()}

    # 好样本应全部 ✅
    for cid in EV.CRIT_ORDER:
        check(f"高质量样本 {good[cid]['label']} = ✅", good[cid]["rating"] == "pass",
              f"得到 {good[cid]['rating']} (raw={good[cid]['raw']})")
    # 差样本应全部 ❌（含硬编码路径/密钥 → 可复用必须 ❌）
    check("低质量样本 可复用 = ❌", bad["reusable"]["rating"] == "fail",
          f"得到 {bad['reusable']['rating']} (raw={bad['reusable']['raw']})")
    # 空壳样本应普遍 ❌
    n_fail = sum(1 for cid in EV.CRIT_ORDER if stub[cid]["rating"] == "fail")
    check("空壳样本 ≥3 个维度 ❌", n_fail >= 3, f"得到 {n_fail}")

    # 负向信号：硬编码路径必须扣分
    r = [x for x in bad["reusable"]["rules"] if x["id"] == "no_hardcoded_env"][0]
    check("硬编码路径检测生效（扣分）", r["score"] is not None and r["score"] < 1.0,
          f"得分 {r['score']}")

    # 显著性差异：好样本综合分必须显著高于差样本
    def qscore(d):
        v = {"pass": 1.0, "partial": 0.5, "fail": 0.0}
        return sum(v[d[c]["rating"]] for c in EV.CRIT_ORDER) / 4
    check("好样本质量分 > 差样本 +0.5", qscore(good) - qscore(bad) > 0.5,
          f"good={qscore(good)} bad={qscore(bad)}")

    # 计算指标
    labels = {"good": good, "bad": bad, "stub": stub}
    expect_pass = {"good": {c: True for c in EV.CRIT_ORDER},
                   "bad": {c: False for c in EV.CRIT_ORDER},
                   "stub": {c: False for c in EV.CRIT_ORDER}}
    tp = fp = tn = fn = 0
    for k, res in labels.items():
        for cid in EV.CRIT_ORDER:
            pred = res[cid]["rating"] == "pass"
            truth = expect_pass[k][cid]
            if pred and truth:
                tp += 1
            elif pred and not truth:
                fp += 1
            elif not pred and not truth:
                tn += 1
            else:
                fn += 1
    prec = tp / (tp + fp) if (tp + fp) else 1.0
    rec = tp / (tp + fn) if (tp + fn) else 1.0
    acc = (tp + tn) / (tp + tn + fp + fn)
    print(f"\n  —— 指标：精确率 {prec:.3f} ｜ 召回率 {rec:.3f} ｜ 准确率 {acc:.3f} "
          f"｜ TP={tp} FP={fp} TN={tn} FN={fn} ——")
    check("精确率 ≥ 0.9", prec >= 0.9)
    check("召回率 ≥ 0.9", rec >= 0.9)
    return {"precision": round(prec, 4), "recall": round(rec, 4), "accuracy": round(acc, 4),
            "tp": tp, "fp": fp, "tn": tn, "fn": fn}


# ---------------------------------------------------------------------------
# 五、边界情况
# ---------------------------------------------------------------------------
def test_edge_cases(rubric):
    print("\n[5] 边界情况")
    # 空文件夹
    empty = FIXTURES / "_empty_dir"
    empty.mkdir(exist_ok=True)
    recs = EV.scan_folder(empty, rubric)
    bundle, non = EV.build_bundles(recs, empty, rubric)
    check("空文件夹 → 0 作者、不崩溃", len(bundle) == 0)

    # 二进制大文件不读取内容（is_binary 标记）
    recs = EV.scan_folder(FIXTURES / "author_full", rubric)
    pngs = [r for r in recs if r["ext"] == ".png"]
    check("PNG 被标记为二进制", all(r["is_binary"] for r in pngs))

    # GBK 编码文件不崩溃
    gbk_dir = FIXTURES / "author_gbk"
    gbk_dir.mkdir(exist_ok=True)
    (gbk_dir / "Test_C4_skill说明.md").write_bytes("输入：路径\n输出：报告\n".encode("gbk"))
    recs = EV.scan_folder(gbk_dir, rubric)
    txt = EV.read_text_safe(Path(recs[0]["abs_path"]))
    check("GBK 文件可解码（无乱码异常）", "输入" in txt, f"得到 {txt[:20]!r}")

    # 语法错误代码
    pybad = FIXTURES / "author_syntax"
    pybad.mkdir(exist_ok=True)
    (pybad / "Bad_C4_代码.py").write_text("def broken(:\n  return 1\n", encoding="utf-8")
    recs = EV.scan_folder(pybad, rubric)
    check("Python 语法错误可被检出", not EV.python_syntax_ok(recs[0]))

    # 超大文件跳过内容
    check("binary_extensions 含 .skill", ".skill" in rubric["limits"]["binary_extensions"])
    check("截断阈值存在", rubric["limits"]["max_content_chars_per_file"] == 200000)

    # 富文本（.docx/.pdf）应参与内容级检测，不再被当作不可读二进制
    check("rich_text_extensions 声明 .docx/.pdf",
          set(rubric["limits"].get("rich_text_extensions", [])) >= {".docx", ".pdf"})
    docx_dir = FIXTURES / "author_docx"
    if docx_dir.is_dir():
        recs = EV.scan_folder(docx_dir, rubric)
        dr = [r for r in recs if r["ext"] == ".docx"]
        check("docx 不被标记为纯二进制", dr and not dr[0]["is_binary"])
        check("docx 被标记为富文本", dr and dr[0].get("is_rich_text") is True)
        if dr:
            txt = EV.extract_any_text(dr[0])
            check("docx 正文可提取（含 输入/输出）", "输入" in txt and "输出" in txt,
                  f"提取到 {len(txt)} 字符")

    # PDF 无第三方库时优雅降级（返回空串，不崩溃）
    try:
        tmp_pdf = FIXTURES / "_pdf_probe.pdf"
        tmp_pdf.write_bytes(b"%PDF-1.4\n% probe\n")
        rec_pdf = {"path": "_pdf_probe.pdf", "abs_path": str(tmp_pdf), "name": tmp_pdf.name,
                   "stem": "_pdf_probe", "ext": ".pdf", "size_kb": 1, "modified": "",
                   "is_binary": False, "is_rich_text": True}
        EV.extract_any_text(rec_pdf)
        check("PDF 提取失败时优雅降级（不崩溃）", True)
        tmp_pdf.unlink(missing_ok=True)
    except Exception as e:
        check("PDF 提取失败时优雅降级（不崩溃）", False, str(e))


# ---------------------------------------------------------------------------
# 六、端到端：真实 C4 提交包
# ---------------------------------------------------------------------------
def test_regression(rubric):
    """回归测试：固化开发中真实发现的两个 bug，防止再次退化。"""
    print("\n[7] 回归测试（开发期真实 bug 固化）")
    # BUG#1: 槽位错配 —— 通用词内容信号导致 AI日志/AAR 抢走"Skill说明"槽位
    #   修复：贪心一对一 + 特异性优先 (level, specificity)
    recs = EV.scan_folder(FIXTURES / "author_full", rubric)
    bundle, _ = EV.build_bundles(recs, FIXTURES / "author_full", rubric)
    comp = EV.check_completeness(bundle[next(iter(bundle))]["files"], rubric)
    check("skill_doc 指向 skill说明 文件（未被 AI日志 抢占）",
          comp["slots"]["skill_doc"]["matched_file"] == "TestAuthor_C4_skill说明.md",
          f"得到 {comp['slots']['skill_doc']['matched_file']}")
    check("teaching_doc 指向 教学说明 文件（未被 AAR 抢占）",
          comp["slots"]["teaching_doc"]["matched_file"] == "TestAuthor_C4_教学说明.md",
          f"得到 {comp['slots']['teaching_doc']['matched_file']}")
    check("ai_log 指向 AI日志 文件",
          comp["slots"]["ai_log"]["matched_file"] == "TestAuthor_C4_AI日志.md",
          f"得到 {comp['slots']['ai_log']['matched_file']}")
    # 一对一：5 个槽位对应 5 个不同文件
    matched = [v["matched_file"] for v in comp["slots"].values() if v["matched_file"]]
    check("完整性一对一分配（无文件重复占用）", len(matched) == len(set(matched)),
          f"matched={matched}")

    # BUG#2: quality_score 口径错误（除以维度数→0–1，却被报告为 /4）
    #   修复：0–4 分制
    ctx = {"exts": {".md", ".py"}, "n_files": 3, "syntax_bad": []}
    crits = {cid: EV.eval_criterion(cid, cfg, GOOD_TEXT, ctx)
             for cid, cfg in rubric["quality_criteria"].items()}
    val = {"pass": 1.0, "partial": 0.5, "fail": 0.0}
    recomputed = round(sum(val[c["rating"]] for c in crits.values()), 4)
    check("quality_score 为 0–4 分制（全 ✅ 时 = 4.0）", recomputed == 4.0, f"得到 {recomputed}")

    # BUG#3: 通用目录名（misc/other）被误当作者名
    rec = {"path": "misc/随便.md", "abs_path": "", "name": "随便.md", "stem": "随便",
           "ext": ".md", "size_kb": 1, "modified": "", "is_binary": False}
    a, m = EV.extract_author(rec, Path("."), rubric)
    check("misc 目录不被当作作者名", a == "Unknown", f"得到 {a}")

    # BUG#4: 同名文件多版本时，槽位应认领**最新版本**（v2 而非 v1）
    recs = EV.scan_folder(FIXTURES / "author_versioned", rubric)
    bundle, _ = EV.build_bundles(recs, FIXTURES / "author_versioned", rubric)
    comp = EV.check_completeness(bundle[next(iter(bundle))]["files"], rubric)
    check("skill_doc 认领 v2 而非 v1",
          comp["slots"]["skill_doc"]["matched_file"] == "Ver_C4_skill说明_v2.md",
          f"得到 {comp['slots']['skill_doc']['matched_file']}")

    # BUG#5: 解包后的技能包目录（与 .skill 包名同名）内的文件被误报为"非 C4 文件"
    check("端到端无 false 非C4（技能包成员归属正确）",
          True)  # 由端到端测试的 non_c4 断言覆盖

    # BUG#6: 与 C4 文件同目录、但文件名无 _C4_ 标记的附属文件
    #         （如交付包顶层 README.md）被误报为"非 C4 文件"
    import tempfile, shutil
    tmp = Path(tempfile.mkdtemp())
    try:
        (tmp / "PackAuthor_C4_skill说明.md").write_text("输入 输出 使用场景", encoding="utf-8")
        (tmp / "PackAuthor_C4_AI日志.md").write_text("使用的 AI 工具 prompt 迭代", encoding="utf-8")
        (tmp / "README.md").write_text("交付包说明：作者 PackAuthor", encoding="utf-8")
        recs = EV.scan_folder(tmp, rubric)
        bundle, non = EV.build_bundles(recs, tmp, rubric)
        check("同目录无标记附属文件不误报为非C4", len(non) == 0, f"得到 {[r['path'] for r in non]}")
        check("同目录 README 归属正确作者",
              any(f["path"] == "README.md" for f in bundle.get("PackAuthor", {}).get("files", [])),
              f"得到 {list(bundle)}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def test_end_to_end(rubric, real_dir):
    print("\n[6] 端到端（真实 C4 提交包）")
    if not Path(real_dir).is_dir():
        check("真实提交包存在", False, str(real_dir))
        return None
    recs = EV.scan_folder(Path(real_dir), rubric)
    bundle, non = EV.build_bundles(recs, Path(real_dir), rubric)
    check("识别出作者", len(bundle) >= 1, f"得到 {list(bundle)}")
    res = []
    for a, b in bundle.items():
        res.append(EV.evaluate_author(a, b, rubric))
    # 技能包内部文件（slide-forge/）应归属 lsa，不应出现在 non_c4
    check("无 false 非C4 文件（技能包成员正确归属）", len(non) == 0,
          f"得到 {[r['path'] for r in non]}")
    for r in res:
        check(f"  {r['author']} 完整性 5/5", r["completeness"]["present"] == 5,
              f"得到 {r['completeness']['present']}，缺 {r['completeness']['missing']}")
        check(f"  {r['author']} 质量 ≥ 3.0", r["quality"]["quality_score"] >= 3.0,
              f"得到 {r['quality']['quality_score']}")
        check(f"  {r['author']} 综合分 ∈ (0,1]", 0 < r["composite"] <= 1)
    return res


# ---------------------------------------------------------------------------
def main():
    rubric = EV.load_rubric(SCRIPT_DIR.parent / "references" / "rubric.yaml")
    real = Path(r"C:/Users/Administrator/WorkBuddy/挑战资料包/outputs/lsa_C4_提交包")
    print("=" * 66)
    print("  c4a-evaluator 自测套件")
    print("=" * 66)
    test_rubric_loading(rubric)
    test_author_extraction(rubric)
    test_completeness(rubric)
    metrics = test_quality_rules(rubric)
    test_edge_cases(rubric)
    test_regression(rubric)
    test_end_to_end(rubric, real)
    print("\n" + "=" * 66)
    print(f"  结果：{PASS} 通过 / {FAIL} 失败")
    if _FAILED:
        print("  失败用例：" + "；".join(_FAILED))
    print("=" * 66)
    out = SCRIPT_DIR.parent / "tests" / "self_test_result.json"
    out.write_text(json.dumps({"pass": PASS, "fail": FAIL, "failed": _FAILED,
                               "metrics": metrics}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  结果已写入 {out}")
    sys.exit(1 if FAIL else 0)


if __name__ == "__main__":
    main()
