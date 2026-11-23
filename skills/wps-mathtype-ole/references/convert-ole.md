# 前台 OLE 转换

## 清单和源文档

在 DOCX 副本的正文普通段落中填写 `[[MT:ID]]`。每个 ID 在当前清单及文档中各出现一次。独立公式用 `display`，占位符独占一段，编号放在清单的 `number` 中。行内公式用 `inline`，省略 `number`。脚本把编号写为段落中的文本，并设置居中及右对齐制表位。

每项提供一段单行 `latex` 或一个包含 MathML 命名空间的 `mathml`。LaTeX 反斜杠按 JSON 规则转义。MathML 使用 Presentation MathML，例如：

```json
{
  "equations": [
    {
      "id": "eq24",
      "kind": "display",
      "latex": "a_{ij}=\\frac{x_i}{x_j}",
      "number": "(24)"
    },
    {
      "id": "inline1",
      "kind": "inline",
      "mathml": "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>&lt;</mo><mn>0</mn></math>"
    }
  ]
}
```

## 运行

```powershell
py -3 scripts/wps_mathtype.py convert-ole --input prepared.docx --manifest batch-01.json --output converted-01.docx
```

脚本会拒绝已有输出文件或带有已打开文档的 WPS COM 会话。运行前保存自己的 WPS 文档。转换时保持桌面解锁，避免切换窗口、使用键鼠或复制内容。脚本会尽力恢复原剪贴板中的纯文本，图片和富文本等其他格式可能被替换。

一个批次共用一份 WPS 工作副本。全部公式转换后保存并验证一次。约 20 条一批便于定位失败。多批时，先准备好全部占位符，每批用单独的 JSON 清单，并以通过验证的上一批输出作为下一批输入。每批使用新输出路径。

## 验收

- 脚本检查 WPS 插入的 `Equation.DSMT4`、DOCX 的内嵌关系与复合文件 CLSID，并将 `Equation Native` 中的原生公式数据与 MathType 编辑器回读的数据核对。
- 检查一直在批次结束后的 DOCX 上进行。失败时查看输出旁的 `*_ole_report.json`，从该批输入重试。报告包含本机路径和公式回读信息，公开分享前应移除或脱敏。
- 在 WPS 中双击代表性的独立和行内公式验证可编辑性，检查编号、排版和数学语义。检查矩阵、分段式和上下标等复杂结构是否完整。
