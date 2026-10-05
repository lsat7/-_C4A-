# C4A 技能提交自动评审（Automated Skill Submission Evaluator）

> **Elite20 挑战 · C4A** —— 把「收发室」升级为「自动阅卷系统」。
> 输入一个本地文件夹路径（WeChat 群同步文件夹 / 手动下载目录），自动识别所有人的 C4 提交，
> 完成**完整性检查 + 四条件质量评审 + 排名 + 改进建议**，输出 Markdown / Excel / HTML / JSON 多格式报告。

---

## 一句话

```
WeChat 群文件夹
      ↓
① 文件采集与识别（按作者分组）
      ↓
② 提交完整性检查（五件必交：skill说明 / 可执行内容 / demo / 教学说明 / AI日志）
      ↓
③ 技能质量评审（C4 四条件：可复用 / 可执行 / 可验证 / IO 明确）
      ↓
④ 评审报告生成（排名 + 建议 + 版本追踪，Markdown + Excel + HTML）
```

---

## 目录结构

```
.
├── README.md                         # 本文件（仓库入口）
├── 00_交付物索引.md                  # 交付物对照表（必交文件 × 挑战条款）
├── lsa_C4A_方案设计.md               # 架构设计 / 评审维度定义 / 技术选型
├── lsa_C4A_评审报告.md               # 真实评审结果（含诚实声明）
├── lsa_C4A_评审报告_Excel详表.xlsx   # 评审 Excel 详表
├── lsa_C4A_评审结果_json.json        # 评审原始 JSON
├── lsa_C4A_评审仪表板.html           # 班级总览仪表板（可视化）
├── lsa_C4A_教学说明.md               # 安装 / 使用 / 输入输出
├── lsa_C4A_AI日志.md                 # 开发全过程 AI 使用记录
├── lsa_C4A_拿来说明.md               # 从 wechat-doc-mapper 拿了什么、改了什么
├── lsa_C4A_AAR.md                    # 事后回顾（额外交付）
├── lsa_C4A_skill-evaluator.skill     # 可安装技能包（ZIP 格式，与 c4a-evaluator/ 同源）
└── c4a-evaluator/                    # 技能源码目录
    ├── SKILL.md                      # 技能入口：工作流 + I/O 契约 + 边界情况
    ├── references/
    │   └── rubric.yaml               # ★ 评分口径唯一来源（阈值/权重/信号，代码零硬编码）
    ├── scripts/
    │   ├── c4a_evaluator.py          # 主程序（Level 1–4 完整流水线，核心纯标准库）
    │   └── self_test.py              # 自测套件（68 项断言 + 精确率/召回率）
    └── tests/
        ├── fixtures/                 # 8 组对抗性夹具（完整/部分/缺失/多版本/语法错误/GBK/docx）
        └── self_test_result.json     # 自测结果
```

---

## 快速开始

### 方式一：解包 `.skill`（推荐）

`.skill` 本质是 ZIP，解包后直接运行：

```bash
# macOS / Linux
unzip lsa_C4A_skill-evaluator.skill -d c4a-evaluator

# Windows（或手动改后缀 .zip 再解压）
```

### 方式二：直接用源码目录

`c4a-evaluator/` 与 `.skill` 包同源，可直接使用。

### 运行评审

```bash
python c4a-evaluator/scripts/c4a_evaluator.py <FOLDER> -o <OUTDIR> --format md,json,excel,html
```

示例：

```bash
python c4a-evaluator/scripts/c4a_evaluator.py "./C4提交文件夹" -o ./report --format md,excel,html
```

> 输入：**一个本地文件夹路径**（挑战关键约束——技能只吃文件夹路径）。
> 输出：按作者分组的评审报告（完整性 ✅/⚠️/❌ + 四条件评分 + 排名 + 改进建议）。

### 运行自测

```bash
python c4a-evaluator/scripts/self_test.py
# 预期：68 通过 / 0 失败；精确率/召回率 = 1.0
```

---

## 技术要点

| 特性 | 说明 |
| --- | --- |
| **拿来主义** | 以 `wechat-doc-mapper.skill` 为基座，复用「扫描 + 作者识别回退链」，重写评审与报告 |
| **单一事实来源** | 所有阈值/权重/信号集中在 `references/rubric.yaml`，代码零硬编码 |
| **规则为主 + LLM 可插拔** | 默认纯规则（确定性、可复现、零依赖），`LLMReviewer` 为可选深审层 |
| **负向信号** | `must_avoid` 规则识别硬伤：硬编码本机路径、`api_key=`、`ast.parse` 语法错误 |
| **混合格式** | 支持 `.md/.docx/.pdf/.skill/.py/.mp4/.png`；`.docx` 可读正文（标准库兜底），`.pdf` 惰性降级 |
| **零依赖核心** | 核心仅用标准库；`openpyxl`/`pyyaml`/`python-docx`/`pypdf` 均为可选增强 |

---

## 评审标准对齐

| 维度 | 权重 | 覆盖 |
| --- | --- | --- |
| 评审准确性 | 30% | 口径外置 + 证据可追溯 + 68 项自测 + 人工抽检 0 误判 |
| 完成级别 | 25% | Level 1–4 全实现 |
| 拿来主义质量 | 20% | 见 `lsa_C4A_拿来说明.md` + 方案设计 §4.6（D1–D4 对照） |
| 可复用性 | 15% | 零依赖核心 + 单文件夹路径输入 + 口径外置 |
| 创新性 | 10% | 一对一贪心完整性判定、`must_avoid` 负向信号、同目录归属、多格式同源报告 |

---

## 作者与许可

- **作者**：lsa
- **挑战**：Elite20 · C4A 技能提交自动评审
- **许可**：MIT
