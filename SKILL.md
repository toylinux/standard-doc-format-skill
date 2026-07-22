---
name: standard-doc-format
description: 生成符合中文标准公文 / 培训感悟排版的 Word 文档。当用户说"使用标准格式"或要求按"标题方正小标宋简体二号居中、署名楷体_GB2312三号居中、正文仿宋_GB2312三号、一级标题黑体、二级标题楷体_GB2312加粗、页边距上下2.54cm左右3.18cm、行距固定30磅、页脚居中纯数字页码"等规范生成 .docx 时使用。
agent_created: true
---

# 标准格式 Word 文档生成技能

## 用途

将用户提供的文档内容（培训感悟、工作总结、公文、心得体会等）按一套固定的中文排版规范生成为 Word（.docx）文件，精确控制字体、字号、对齐、缩进、页边距、行距与页码。

## 触发条件

- 用户说"使用标准格式""按标准格式""用标准排版"等。
- 用户要求生成符合以下任一项特征的 docx：标题方正小标宋简体二号居中、署名楷体_GB2312 居中、正文仿宋_GB2312 三号、一级标题黑体、二级标题楷体加粗、页边距 2.54/3.18cm、行距固定 30 磅、页脚居中纯数字页码。

## 格式规范（务必逐项落实）

| 元素 | 字体 | 字号 | 对齐 / 缩进 |
|------|------|------|------|
| 标题（第一行） | 方正小标宋简体 | 二号 (22pt) | 居中，无缩进 |
| 署名（第二行） | 楷体_GB2312 | 三号 (16pt) | 居中，无缩进 |
| 一级标题 | 黑体 | 三号 (16pt) | 左对齐，首行缩进 2 字符 |
| 二级标题 | 楷体_GB2312 | 三号 (16pt) 加粗 | 左对齐，首行缩进 2 字符 |
| 正文 | 仿宋_GB2312 | 三号 (16pt) | 两端对齐，首行缩进 2 字符 |

- 页边距：上 / 下 2.54cm，左 / 右 3.18cm。
- 行距：固定值 30 磅（全篇统一）。
- 页码：页脚居中，纯数字（PAGE 域自动更新），字号小四 (12pt)。
- 署名可选；缺省时跳过署名行。

## 环境准备（一次性）

脚本依赖 `python-docx`。使用受管 Python 环境，在隔离 venv 中安装（不要污染用户环境）：

```bash
C:\Users\toy\.workbuddy\binaries\python\versions\3.13.12\python.exe -m venv C:\Users\toy\.workbuddy\binaries\python\envs\default
C:\Users\toy\.workbuddy\binaries\python\envs\default/Scripts/pip.exe install python-docx
```

运行脚本所用的解释器即上述 venv 中的
`C:\Users\toy\.workbuddy\binaries\python\envs\default/Scripts/python.exe`。

## 使用流程

1. 整理用户的内容，写成带标签的纯文本输入文件，保存到当前工作目录（如 `doc_input.txt`）：

   ```
   [TITLE] 如何做一名"四有"好教师
   [SIGN] 古塔区保二小学  李晓莉
   [BODY] （引言段落……）
   [H1] 一、有理想信念，做学生成长路上的引路人
   [BODY] （正文……）
   [H2] （一）立足学科特点，在真实情境中激发兴趣
   [BODY] （正文……）
   [H1] 二、有道德情操，用言行诠释师德风范
   [BODY] （正文……）
   [H1] 三、有扎实学识，立足英语课堂锤炼业务能力
   [BODY] （正文……）
   [H1] 四、有仁爱之心，温暖每一个孩子的童年
   [BODY] （正文……）
   [BODY] （结语……）
   ```

   - 标签大小写均可，行内标签后紧跟文本。
   - 无标签的普通非空行按正文处理。
   - 一个段落占一行；长段落不要换行（换行会被当成两个段落）。

2. 调用脚本生成 docx（输出到当前工作目录，便于直接预览/下载）：

   ```bash
   C:\Users\toy\.workbuddy\binaries\python\envs\default/Scripts/python.exe "<skill_dir>/scripts/standard_doc.py" doc_input.txt 标准格式文档.docx
   ```

   其中 `<skill_dir>` 为本技能所在目录（通常为
   `C:\Users\toy\.workbuddy\skills\standard-doc-format`）。

3. 用 `present_files` 将生成的 .docx 呈现给用户。

## 注意事项

- `方正小标宋简体` 为方正字库字体，若本机未安装，Word 会回退为默认字体；其余楷体_GB2312、仿宋_GB2312、黑体为 Windows 常用字体，一般已自带。如用户要求保证标题字体，提示其安装对应字体即可。
- 生成后建议用 python-docx 回读校验：标题 eastAsia=方正小标宋简体/22pt/居中，正文 eastAsia=仿宋_GB2312/16pt，一级 eastAsia=黑体，二级 eastAsia=楷体_GB2312 且 bold，页脚含 PAGE 域且居中，页边距 2.54/3.18cm，行距 EXACTLY=30pt。
- 若用户临时要求调整某项（如页码改"第 X 页"、标题不缩进等），直接改输入或脚本参数后重新生成，不必改动本技能核心规范。
