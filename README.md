# WPS MathType OLE

这是一个面向 Windows 桌面版 WPS 文字的 Codex 技能。它把 LaTeX 或 Presentation MathML 插入 DOCX，生成可在桌面 MathType 中双击编辑的 `Equation.DSMT4` 公式。支持行内公式、独立公式居中和编号右对齐。

## 安装到 Codex

下载或克隆本仓库，将 [`skills/wps-mathtype-ole`](skills/wps-mathtype-ole/SKILL.md) 整个文件夹复制到 `%USERPROFILE%\.agents\skills\`。如果已有同名技能，请先把旧版本移出技能目录，再复制新版。在 Codex 中输入 `$wps-mathtype-ole`。如果技能列表未刷新，重启 Codex。

也可在 Codex 中用 `$skill-installer` 安装。把下面的 `YOUR_USERNAME` 换成仓库所有者的 GitHub 用户名：

```text
$skill-installer 请安装 https://github.com/YOUR_USERNAME/wps-mathtype-ole/tree/main/skills/wps-mathtype-ole
```

参见 [OpenAI 技能安装文档](https://learn.chatgpt.com/docs/build-skills)。

## 运行环境

- Windows 桌面版 WPS 文字，支持 `kwps.application` COM 自动化。
- 桌面版 MathType 7 for Windows，能在 WPS 中嵌入并编辑传统 OLE 公式。Wiris 从 MathType 7.11.1 起提供 64 位 WPS 兼容支持。[MathType 版本说明](https://docs.wiris.com/en_US/mathtype-release-notes/mathtype-office-tools-release-notes)
- Windows Python 3.10 或以上及 [Python 依赖](skills/wps-mathtype-ole/requirements.txt)。`convert-ole` 需要已解锁的交互式桌面。

## 快速开始

先保存并退出 WPS 文字。打开 PowerShell，从仓库根目录安装依赖并运行预检：

```powershell
Set-Location .\skills\wps-mathtype-ole
py -3 -m pip install -r requirements.txt
py -3 .\scripts\wps_mathtype.py self-check
```

预检成功后，在同一 PowerShell 窗口运行示例。示例会生成一个含三条公式占位符的 DOCX，并把源文件、输出和转换报告放在 `work` 下。

```powershell
py -3 .\scripts\create_demo.py --output ..\..\work\demo_source.docx
py -3 .\scripts\wps_mathtype.py convert-ole --input ..\..\work\demo_source.docx --manifest .\examples\manifest.json --output ..\..\work\demo_mathtype.docx
py -3 .\scripts\wps_mathtype.py inspect --input ..\..\work\demo_mathtype.docx
```

`py` 是 Windows Python 启动器。若系统使用其他入口，可改用已确认指向 Windows Python 解释器的 `python`。转换期间让 WPS 和 MathType 保持可见，避免切换窗口、输入或使用剪贴板。打开输出，在 WPS 中双击示例公式确认可编辑，再检查编号和对齐。

## 转换自己的文档

在 DOCX 副本中把目标位置改为唯一的 `[[MT:ID]]`，并按 [示例清单](skills/wps-mathtype-ole/examples/manifest.json) 提供对应的表达式。已有 Word 或 WPS 原生公式需要先提取为 TeX 或 Presentation MathML，再准备占位符与清单。转换命令按清单工作。

`convert-ole` 可以在一个 WPS 工作副本中逐条转换。大批量时建议每批约 20 条，并用上一批通过验证的 DOCX 作为下一批的输入。若本机已验证 WPS 能通过 COM 调用 MathType 的 Toggle TeX，可使用 `convert` 宏入口。

详见 [技能说明](skills/wps-mathtype-ole/SKILL.md)、[清单格式与分批流程](skills/wps-mathtype-ole/references/convert-ole.md) 和 [宏模式与兼容性](skills/wps-mathtype-ole/references/formatting-and-compat.md)。
