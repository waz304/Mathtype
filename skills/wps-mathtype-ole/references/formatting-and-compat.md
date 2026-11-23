# 宏模式与兼容性

## Toggle TeX 宏模式

此模式依赖本机 WPS 的 MathType 加载项。先在测试文档中手动验证 Toggle TeX 有效，再在保存并退出 WPS 文字后，通过 COM 对测试副本运行：

```powershell
py -3 scripts/wps_mathtype.py convert --input prepared.docx --manifest manifest.json --output converted.docx
```

宏模式接受 LaTeX 清单，默认调用 `MTCommand_TeXToggle`。如本机验证了其他宏名，可用 `--macro` 指定。WPS 的 `Application.Run` 文档描述的 JS 宏支持不代表传统 MathType VBA 宏可用。宏调用失败时查看 `*_mathtype_report.json`，改用 [前台 OLE 转换](convert-ole.md)。[WPS Application.Run 文档](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wps/Application/member/Run)

## 检查兼容性和排版

- 预检报告 `NEEDS_WPS_CLOSE` 时，保存并退出 WPS 文字后重试。检查 `kwps.application` 可用性与 `Equation.DSMT4` 注册情况。注册缺失时检查桌面 MathType 及其 WPS 集成。[WPS 应用对象文档](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wps/Application/obj) · [MathType WPS 指南](https://mathtype.org/blog/mathtype-for-wps.html)
- 设定 MathType 中的字体和字号后再检查正文基线。独立公式使用居中制表位，编号使用右对齐制表位。
- 若 TeX 翻译不完整，使用一条合成公式测试当前安装。在支持的情况下，改用表达相同数学内容的 Presentation MathML。WPS 与 MathType 的安装版本和编辑器语言会影响前台路径。
