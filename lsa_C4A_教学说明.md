# lsa_C4A_教学说明.md

> 怎么安装、怎么用、输入什么得到什么。目标是**让一位完全没接触过的老师/助教，5 分钟内跑出第一份评审报告**。

---

## 一、这个技能能帮你做什么

一句话：**给一个文件夹路径，自动生成全班 C4 提交的评审报告（谁交了 / 交齐没 / 好不好 / 怎么改）。**

适合场景：
- 你是老师/助教，群里收了一堆 C4 提交，想快速了解全班情况；
- 你想在提交前**自检**自己的 C4 技能包是否满足五件必交 + 四条件；
- 你想把评审结果发到群里，帮同学发现问题。

---

## 二、安装（30 秒）

### 方式 A：直接用技能包（推荐）

```
1. 拿到 lsa_C4A_skill-evaluator.skill（本质是 ZIP，可解压）
2. 解压：把 .skill 改名为 .zip 后解压，或用命令
   （macOS/Linux）: unzip lsa_C4A_skill-evaluator.skill
   （Windows）    : 复制一份改名为 .zip，右键解压
3. 得到 c4a-evaluator/ 目录
```

### 方式 B：直接复制源码目录

把 `c4a-evaluator/` 整个目录拷到任意位置即可——**没有任何需要"安装"的东西**。

### 环境要求

| 依赖 | 是否必须 | 说明 |
| --- | --- | --- |
| **Python 3.8+** | ✅ 必须 | 核心只用标准库 |
| pyyaml | 可选 | 装了更稳；没装自动降级内置解析器 |
| openpyxl | 可选 | 装了才能输出 Excel；没装跳过 Excel |

> 检查 Python：`python --version` 或 `python3 --version`。
> 安装可选依赖（需要时才装）：
> ```bash
> pip install pyyaml openpyxl
> ```

---

## 三、上手（3 步）

### Step 1 — 把 C4 提交放进一个文件夹

无论文件从哪来（WeChat 手动下载 / 群文件导出 / 微信 PC 端缓存目录），
**放进一个文件夹即可**。技能不关心文件怎么来的。

```
我的C4提交/
├── lsa_C4_skill说明.md
├── lsa_C4_slide-forge.skill
├── lsa_C4_demo.png
├── lsa_C4_教学说明.md
└── lsa_C4_AI日志.md
```

> 💡 文件命名建议遵守 `姓名拼音_C4_内容描述.扩展名`。**不遵守也能跑**——技能会回退到
> 子目录名/文档内容识别作者，但规范命名识别率更高。

### Step 2 — 跑一行命令

```bash
cd c4a-evaluator

# 最简：只出 Markdown + JSON
python scripts/c4a_evaluator.py "我的C4提交" -o ./评审输出

# 全格式：Markdown + JSON + Excel + HTML 仪表板
python scripts/c4a_evaluator.py "我的C4提交" -o ./评审输出 --format md,json,excel,html
```

Windows 路径记得加引号：`python scripts\c4a_evaluator.py "D:\群里文件" -o out`

### Step 3 — 看结果

命令会先在终端打印一个速览：

```
=== 评审完成：识别 5 位作者 / 30 个文件 ===
  lsa       完整性 5/5  质量 4.0/4  综合 1.0   ✅✅✅✅
  zhangwei  完整性 5/5  质量 4.0/4  综合 1.0   ✅✅✅✅
  chenyu    完整性 4/5  质量 4.0/4  综合 0.92  ✅✅✅✅
  liqiang   完整性 3/5  质量 2.0/4  综合 0.54  ❌✅❌✅
  zhaolei   完整性 2/5  质量 2.5/4  综合 0.535 ⚠️✅❌✅
```

输出目录里的文件：

| 文件 | 用途 |
| --- | --- |
| `c4a_eval_report.md` | **主报告**：班级总览 + 排名 + 每位作者详情 + 改进建议（可直接贴群） |
| `c4a_eval_result.json` | 结构化结果（供二次处理） |
| `c4a_eval_report.xlsx` | Excel 详表（总览/明细/文件清单三个 Sheet） |
| `c4a_eval_dashboard.html` | 可视化仪表板（双击用浏览器打开） |

---

## 四、输入 / 输出速查

| | 内容 |
| --- | --- |
| **输入** | 一个本地文件夹路径（唯一必填参数） |
| **输出** | 评审报告（Markdown）+ 结构化结果（JSON）+ 可选 Excel、HTML |

### 命令行参数

| 参数 | 作用 | 默认 |
| --- | --- | --- |
| `<folder>` | 待评审文件夹（**必填**） | — |
| `-o / --outdir` | 输出目录 | `./c4a_eval_out` |
| `--format` | 输出格式，逗号分隔：`md,json,excel,html` | `md,json` |
| `--rubric` | 自定义评分口径 YAML 路径 | 内置 `references/rubric.yaml` |
| `--llm` | LLM 深审层：`none` | `none`（纯规则） |

---

## 五、评分怎么算（看得懂才信得过）

```
完整性分 = 命中文件数 / 5
质量分   = 四条件之和（每条件 ✅=1 / ⚠️=0.5 / ❌=0，满分 4）
综合分   = 完整性分 × 0.40 + (质量分 ÷ 4) × 0.60
```

四条件的判定规则全部写在 `references/rubric.yaml`，每条都附**证据**（命中/未命中哪些信号）。
报告里每位作者都有"展开逐规则证据"，点击即可看到争议点。

---

## 六、常见坑（8 条，都来自真实踩坑）

1. **槽位错配**：早期版本会让"含输入/输出字样的任何文档"都满足"Skill 说明"槽位，
   导致假齐全。→ 已改为**一对一贪心分配 + 特异性优先**（文件名信号长的优先）。
2. **多版本选错**：同一作者有 `_v1`/`_v2` 时，槽位必须认领**最新版**。→ 已按版本号排序认领。
3. **解包技能包被误判为非 C4**：与 `.skill` 同名的子目录（如 `slide-forge/`）内的文件
   应归属该作者，而不是报"非 C4 文件"。→ 已按包名吸收。
4. **编码问题**：中文文件名的 GBK 文件读成乱码。→ 已按 UTF-8→GBK→GB18030→latin-1 回退。
5. **Windows 路径引号**：路径含空格/中文时命令要加引号，否则参数被截断。
6. **没装 openpyxl**：`--format excel` 会静默跳过并提示，不报错——先 `pip install openpyxl`。
7. **没装 pyyaml**：会自动用内置解析器，**不影响结果**（自测已验证与 pyyaml 一致）。
8. **改了口径没复测**：改 `rubric.yaml` 后务必跑 `python scripts/self_test.py`，
   防止改坏判定逻辑（61 项断言会立刻告诉你）。

---

## 七、优化技巧（进阶）

| 目标 | 做法 |
| --- | --- |
| **换课程/换标准** | 只改 `references/rubric.yaml`（阈值/权重/信号），**代码零改动** |
| **调整完整性 vs 质量权重** | 改 `rubric.yaml` 的 `composite.completeness_weight` / `quality_weight` |
| **四条件不等权** | 改 `criteria_weights`（如把"可执行"提到 0.4） |
| **接入 LLM 深审** | 实现 `LLMReviewer.review(text, crits)` 并注入 `run(..., llm=...)` |
| **嵌入自己的流程** | 直接 `import c4a_evaluator as EV; EV.run(folder, rubric, out, formats)` |
| **验证改动的正确性** | `python scripts/self_test.py`（61 项断言 + 精确率/召回率输出） |

---

## 八、自检清单（提交 C4 前自己跑一遍）

```bash
python scripts/c4a_evaluator.py "你的提交文件夹" -o ./自检
```

然后逐项核对：

- [ ] 完整性是不是 5/5？（缺哪个补哪个）
- [ ] 可复用是不是 ✅？（有没有硬编码本机路径？有没有安装说明？环境要求写了吗？）
- [ ] 可执行是不是 ✅？（代码能跑吗？`.skill` 包结构完整吗？）
- [ ] 可验证是不是 ✅？（有示例吗？写了预期结果吗？）
- [ ] IO 明确是不是 ✅？（有没有"输入…输出…"一句话？）

---

## 九、复现本文档的所有结果

```bash
cd c4a-evaluator
python scripts/self_test.py                    # 61 项自测
python scripts/c4a_evaluator.py "<任意C4文件夹>" -o out --format md,json,excel,html
```
