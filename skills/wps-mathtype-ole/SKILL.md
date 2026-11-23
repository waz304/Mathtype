---
name: wps-mathtype-ole
description: 在 Windows 桌面版 WPS 文字中把 DOCX 指定位置的 LaTeX 或 Presentation MathML 转成可双击编辑的 MathType OLE 公式。用于行内公式、带编号的独立公式及批量替换，要求 WPS COM 和已注册的桌面 MathType。
---

# WPS 文字中的 MathType OLE 公式

## 环境

- 使用 Windows 桌面版 WPS 文字的 `kwps.application` 和可编辑 `Equation.DSMT4` 对象的桌面版 MathType。64 位 WPS 使用 MathType 7.11.1 或更新版。安装 [Python 依赖](requirements.txt)。
- 从本技能目录运行 `py -3 scripts/wps_mathtype.py self-check`。就绪状态只表明 COM 和 OLE 预检通过。先用生成的示例或一条合成公式测试实际插入、保存和双击编辑。
- 先保存并关闭正在编辑的 WPS 文档。前台转换需要解锁的桌面，并在运行期间使用 WPS、MathType 和剪贴板。

## 处理文档

1. 保留原稿，在 DOCX 副本的正文普通段落中放置唯一的 `[[MT:ID]]`。独立公式的占位符单独占一段，行内公式保留在原句中。转换现有原生公式时，先提取并核对数学内容，再替换目标位置。记录编号。遇到自动编号、书签或交叉引用时，检查替换后引用的目标。
2. 写一份与占位符逐一对应的 UTF-8 JSON [公式清单](examples/manifest.json)。`convert-ole` 接受 LaTeX 和 Presentation MathML。矩阵、对齐式或分段式可优先试用保留结构的 MathML。
3. 用 `py -3 scripts/wps_mathtype.py convert-ole --input prepared.docx --manifest manifest.json --output converted.docx` 输出新文件。批次使用、剪贴板影响及校验见 [前台 OLE 转换](references/convert-ole.md)。只有在本机验证 Toggle TeX 可由 WPS COM 调用时才选用 [宏模式](references/formatting-and-compat.md)。
4. 前台模式校验实际嵌入的 MathType OLE 类和公式原生数据。使用 `py -3 scripts/wps_mathtype.py inspect --input converted.docx` 辅助检查。验收时在 WPS 双击代表性公式确认可编辑，核对表达式、行内基线、独立公式居中和编号右对齐。转换失败时查看相应的 JSON 报告。
