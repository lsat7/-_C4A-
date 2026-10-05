# TestAuthor_C4_skill说明.md — demo-skill 技能说明

## 解决什么问题
把一批零散文件整理成结构化清单，省去手工归类。

## 使用场景
通用，适用于任何目录；兼容 Windows / macOS / Linux。

## 输入输出
- **输入**：一个文件夹路径（本地目录）。
- **输出**：一份 Markdown 清单文件。
一句话：输入文件夹路径，输出结构化清单。

## 安装与使用步骤
1. 环境要求：python 3.8 以上，零依赖（仅标准库），无需 pip install。
2. 运行：`python demo_skill.py <文件夹路径> -o 清单.md`

## 示例
```
$ python demo_skill.py ./samples -o out.md
结果：成功，✓ 通过，共 12 条
```
预期输出：out.md 中包含 12 行清单。
