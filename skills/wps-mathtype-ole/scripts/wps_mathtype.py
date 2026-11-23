#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Insert and verify MathType OLE equations in WPS Writer.

Run in Windows desktop with WPS Writer and MathType for Windows.
Unverified WPS integrations fail closed: never substitute OMML/images for MathType.
"""

from __future__ import annotations

import argparse
import csv
import io
import subprocess
import json
import os
import re
import shutil
import sys
import tempfile
import traceback
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

SKILL_NAME = "wps-mathtype-ole"
MARKER_RE = re.compile(r"\[\[MT:([A-Za-z0-9._-]+)\]\]")
ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,79}$")
# Native WPS COM constants (avoid dependence on generated type libraries).
WD_FIND_STOP = 0
WD_TAB_CENTER = 1
WD_TAB_RIGHT = 2
WD_TAB_LEADER_SPACES = 0
WD_ALIGN_PARAGRAPH_LEFT = 0


class MathTypeError(RuntimeError):
    """Failure in MathType conversion or verification."""


def _json_print(value):
    print(json.dumps(value, ensure_ascii=False, indent=2))


def _win32_client():
    if sys.platform != "win32":
        raise MathTypeError("只能在 Windows 桌面 Python 中调用 WPS/MathType COM；Linux/WSL/macOS 无法执行。")
    try:
        import pythoncom
        import win32com.client
        return pythoncom, win32com.client
    except ImportError as exc:
        raise MathTypeError("缺少 pywin32。请在 Windows 上执行：python -m pip install pywin32") from exc


def _ole_classes(doc):
    """Collect OLE object classes. A report alone never substitutes for user visual QA."""
    results = []
    for container_name in ("InlineShapes", "Shapes"):
        container = getattr(doc, container_name)
        for index in range(1, container.Count + 1):
            shape = container.Item(index)
            try:
                cls = str(shape.OLEFormat.ClassType).strip()
            except Exception:
                continue
            if cls:
                results.append({"container": container_name, "index": index, "class": cls})
    return results


def _is_mathtype_class(value):
    name = value.strip().lower()
    # MathType's classic OLE ProgID. Do not accept Equation.3/Word.Equation as equivalent.
    return name.startswith("equation.dsmt") or name.startswith("mathtype.")


def _math_objects(doc):
    return [item for item in _ole_classes(doc) if _is_mathtype_class(item["class"])]


def _wps_writer_running():
    """Conservative check: refuse to touch an already-open WPS Writer session."""
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq wps.exe", "/FO", "CSV", "/NH"],
            check=True, capture_output=True, text=True, encoding="mbcs", timeout=10,
        )
        return any(
            row and row[0].strip().lower() == "wps.exe"
            for row in csv.reader(io.StringIO(result.stdout))
        )
    except Exception as exc:
        raise MathTypeError(f"无法可靠检测是否有 WPS 文字正在运行，请手动关闭 WPS 并重试：{exc}") from exc


def _start_wps():
    if _wps_writer_running():
        raise MathTypeError("检测到 WPS 文字进程 wps.exe。先保存并退出所有 WPS 文字窗口，避免操作或关闭已有文件。")
    pythoncom, win32 = _win32_client()
    pythoncom.CoInitialize()
    try:
        # WPS Open Platform documents this specific ProgID; not Word.Application.
        app = win32.DispatchEx("kwps.application")
        app.Visible = True
        try:
            app.DisplayAlerts = 0
        except Exception:
            pass
        return pythoncom, app
    except Exception as exc:
        pythoncom.CoUninitialize()
        raise MathTypeError(
            "无法启动 WPS 文字 COM (kwps.application)。请确认已安装 Windows 桌面版 WPS、"
            "对应版本提供 COM 自动化，并检查 Python 的 pywin32 环境。原始错误：" + str(exc)
        ) from exc


def self_check():
    report = {
        "platform": sys.platform,
        "python": sys.version.split()[0],
        "host": "WPS Writer (kwps.application)",
        "requirements": "MathType for Windows 7.11.1+ recommended for 64-bit WPS",
    }
    if sys.platform != "win32":
        report["status"] = "BLOCKED"
        report["message"] = "只有 Windows 本地 WPS Writer + MathType 传统插件可测试经典 OLE 插入。"
        _json_print(report)
        return 2
    try:
        import winreg
        clsid = None
        for candidate in ("Equation.DSMT4", "Equation.DSMT3"):
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, candidate + "\\CLSID") as key:
                    clsid, _ = winreg.QueryValueEx(key, "")
                    report["registered_ole_class"] = candidate
                    break
            except OSError:
                pass
        report["ole_class_registered"] = bool(clsid)
        report["wps_writer_already_running"] = _wps_writer_running()
        if report["wps_writer_already_running"]:
            report["status"] = "NEEDS_WPS_CLOSE"
            report["message"] = "请先保存并退出 WPS 文字窗口，再运行 self-check。不会接管已有文档。"
            _json_print(report)
            return 2
        pythoncom, app = _start_wps()
        try:
            report["wps_version"] = str(app.Version)
            try:
                report["wps_name"] = str(app.Name)
            except Exception:
                pass
            loaded = []
            try:
                for i in range(1, app.AddIns.Count + 1):
                    addin = app.AddIns.Item(i)
                    name = str(addin.Name)
                    if "mathtype" in name.lower():
                        loaded.append({"name": name, "installed": bool(addin.Installed)})
            except Exception as exc:
                report["addins_probe_warning"] = str(exc)
            report["mathtype_wps_addins_seen"] = loaded
        finally:
            app.Quit()
            pythoncom.CoUninitialize()
        report["status"] = "READY_FOR_REAL_CONVERSION_TEST" if clsid else "BLOCKED"
        report["message"] = (
            "WPS COM 可启动，已找到传统 MathType OLE 类。请用示例 DOCX 测试所选转换模式，"
            "并双击验证结果。convert-ole 需要 Equation.DSMT4；convert 需要可调用的 Toggle TeX 宏。"
            if clsid else "未发现 MathType 传统 OLE 类；检查是否装有 MathType for Windows 桌面版。"
        )
        _json_print(report)
        return 0 if clsid else 2
    except Exception as exc:
        report["status"] = "BLOCKED"
        report["error"] = str(exc)
        _json_print(report)
        return 2


def _load_manifest(path, *, allow_mathml=False):
    raw = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    rows = raw.get("equations") if isinstance(raw, dict) else None
    if not isinstance(rows, list) or not rows:
        raise MathTypeError("manifest 必须是含有非空 equations 数组的 JSON。")
    seen = set()
    results = []
    for index, row in enumerate(rows, 1):
        if not isinstance(row, dict):
            raise MathTypeError(f"第 {index} 条不是对象。")
        ident = row.get("id")
        kind = row.get("kind")
        latex = row.get("latex")
        number = row.get("number", "")
        if not isinstance(ident, str) or not ID_RE.fullmatch(ident):
            raise MathTypeError(f"第 {index} 条 id 无效。仅允许英数字、_、-、.；长度不超过80。")
        if ident in seen:
            raise MathTypeError(f"重复的公式 id：{ident}")
        seen.add(ident)
        if kind not in ("inline", "display"):
            raise MathTypeError(f"{ident}: kind 必须为 inline 或 display。")
        if allow_mathml and "mathml" in row:
            markup = row["mathml"]
            if "latex" in row or not isinstance(markup, str) or not markup.strip():
                raise MathTypeError(f"{ident}: 请仅提供一项非空 mathml 或 latex。")
            content = {"mathml": markup.strip()}
        else:
            if not isinstance(latex, str) or not latex.strip() or "\r" in latex or "\n" in latex:
                raise MathTypeError(f"{ident}: latex 必须是非空单行 TeX。")
            if any(ord(ch) >= 0x10000 for ch in latex):
                raise MathTypeError(f"{ident}: 不支持 TeX 内的非 BMP 字符（避免 WPS UTF-16 Range 偏移）。")
            content = {"latex": latex.strip()}
        if not isinstance(number, str) or len(number) > 40 or "\n" in number or "\r" in number:
            raise MathTypeError(f"{ident}: number 不是有效的短文本。")
        if kind == "inline" and number:
            raise MathTypeError(f"{ident}: 行内公式不得设置 number。")
        results.append({"id": ident, "kind": kind, "number": number, **content})
    return results


def _find_marker(doc, marker):
    entire = str(doc.Content.Text)
    count = entire.count(marker)
    if count != 1:
        raise MathTypeError(f"占位符 {marker} 在 WPS 文档主体中出现 {count} 次；必须恰好一次（不处理文本框/页眉）。")
    rng = doc.Content.Duplicate
    find = rng.Find
    find.ClearFormatting()
    find.Text = marker
    find.MatchCase = True
    find.MatchWildcards = False
    find.Forward = True
    find.Wrap = WD_FIND_STOP
    if not find.Execute():
        raise MathTypeError(f"WPS COM 未能定位 {marker}")
    return rng


def _format_display_paragraph(doc, paragraph_start):
    """Center the equation using a center tab; right-align its number using right tab."""
    para = doc.Range(paragraph_start, paragraph_start).Paragraphs.Item(1)
    fmt = para.Format
    fmt.Alignment = WD_ALIGN_PARAGRAPH_LEFT
    fmt.LeftIndent = 0
    fmt.RightIndent = 0
    fmt.FirstLineIndent = 0
    fmt.TabStops.ClearAll()
    try:
        ps = para.Range.Sections.Item(1).PageSetup
    except Exception:
        ps = doc.PageSetup
    width = float(ps.PageWidth - ps.LeftMargin - ps.RightMargin)
    if width < 80:
        raise MathTypeError("WPS 当前节的版心宽度异常；未设置公式制表位。")
    fmt.TabStops.Add(Position=width / 2.0, Alignment=WD_TAB_CENTER, Leader=WD_TAB_LEADER_SPACES)
    fmt.TabStops.Add(Position=width - 3.0, Alignment=WD_TAB_RIGHT, Leader=WD_TAB_LEADER_SPACES)


def _insert_one(app, doc, row, macro):
    marker = "[[MT:" + row["id"] + "]]"
    rng = _find_marker(doc, marker)
    pos = int(rng.Start)
    kind = row["kind"]
    if kind == "display":
        para_text = str(rng.Paragraphs.Item(1).Range.Text)
        # The paragraph should contain only this marker, no old numbering or text.
        if para_text.strip("\r\x07 \t") != marker:
            raise MathTypeError(f"{row['id']}: 独立公式占位符必须独占一个 WPS 段落；不得与既有编号/正文混写。")
        tex = "\\[" + row["latex"] + "\\]"
        replacement = "\t" + tex + ("\t" + row["number"] if row["number"] else "")
        offset = 1
    else:
        tex = "$" + row["latex"] + "$"
        replacement = tex
        offset = 0

    before = _math_objects(doc)
    rng.Text = replacement
    if kind == "display":
        _format_display_paragraph(doc, pos)
    # Position by WPS Range, never SendKeys or clipboard; parser sees only the selected TeX.
    eqrange = doc.Range(pos + offset, pos + offset + len(tex))
    doc.Activate()
    eqrange.Select()
    try:
        app.Run(macro)
    except Exception as exc:
        raise MathTypeError(
            f"{row['id']}: MathType WPS 宏 {macro} 调用失败。"
            "请检查 MathType for Windows WPS 加载项、WPS 宏策略与插件版本；不能回退为 OMML。"
            f" 原始错误：{exc}"
        ) from exc
    after = _math_objects(doc)
    if len(after) != len(before) + 1:
        all_ole = _ole_classes(doc)
        raise MathTypeError(
            f"{row['id']}: 调用宏后 MathType OLE 个数从 {len(before)} 变为 {len(after)}"
            "（应恰好增加1）。可能未转换或生成了 WPS 原生公式。"
            f" 已检测 OLE 类：{all_ole[-6:]}"
        )
    if marker in str(doc.Content.Text):
        raise MathTypeError(f"{row['id']}: 占位符仍然存在。")
    if kind == "display":
        # The MathType macro can change paragraph style; reapply our tab stop layout.
        _format_display_paragraph(doc, pos)
    newly_seen = after[-1]["class"]
    return {"id": row["id"], "kind": kind, "number": row["number"], "status": "created-mathtype-ole", "ole_class": newly_seen}


def inspect_docx(path):
    p = Path(path)
    if not p.exists():
        raise MathTypeError(f"文件不存在：{p}")
    with zipfile.ZipFile(p, "r") as z:
        names = z.namelist()
        if "word/document.xml" not in names:
            raise MathTypeError("不是有效的 .docx 主体结构。")
        text = z.read("word/document.xml")
        root = ET.fromstring(text)
        omath = sum(1 for elm in root.iter() if elm.tag == "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath")
        # This is an auxiliary signature, NOT a reliable class / editability check.
        ole_parts = [n for n in names if n.startswith("word/embeddings/") and n.lower().endswith(".bin")]
        relationship_xml = z.read("word/_rels/document.xml.rels") if "word/_rels/document.xml.rels" in names else b""
        ole_rels = relationship_xml.count(b"/oleObject")
        ole_progids = []
        for elem in root.iter():
            if elem.tag.endswith("}OLEObject"):
                progid = elem.attrib.get("ProgID", "")
                if progid:
                    ole_progids.append(progid)
        error_previews = []
        for name in names:
            if name.startswith("word/media/") and name.lower().endswith((".wmf", ".emf")):
                blob = z.read(name)
                if b"Tex translation failed" in blob or b"translation failed" in blob:
                    error_previews.append(name)
        return {
            "file": str(p.resolve()),
            "embedded_ole_bin_parts": len(ole_parts),
            "ole_relationships": ole_rels,
            "ole_object_progids": ole_progids,
            "mathtype_progids": [s for s in ole_progids if _is_mathtype_class(s)],
            "word_native_omath_elements": omath,
            "tex_translation_failed_previews": error_previews,
            "caution": "DOCX XML 是辅助检查。对确切可编辑性仍应通过 WPS COM 的 OLEFormat.ClassType 与 WPS 双击验证；红色错误字需要人工检查。",
        }


def convert(args):
    src = Path(args.input).expanduser().resolve()
    dst = Path(args.output).expanduser().resolve()
    if not src.is_file() or src.suffix.lower() != ".docx":
        raise MathTypeError("输入文件必须是存在的 .docx。")
    if src == dst:
        raise MathTypeError("输出文件不得与源文档同名同路径。")
    if dst.suffix.lower() != ".docx":
        raise MathTypeError("输出必须使用 .docx 扩展名。")
    if dst.exists() and not args.overwrite:
        raise MathTypeError("输出文件已存在；拒绝覆盖。需要覆盖请明确加 --overwrite。")
    manifest = _load_manifest(args.manifest)
    dst.parent.mkdir(parents=True, exist_ok=True)
    report_path = dst.with_name(dst.stem + "_mathtype_report.json")
    report = {
        "source": str(src), "requested": len(manifest), "converted": 0,
        "status": "in_progress", "equations": [], "macro": args.macro,
        "output": str(dst),
    }
    pythoncom = None
    app = None
    doc = None
    with tempfile.TemporaryDirectory(prefix="mt_", dir=str(dst.parent)) as tmp:
        tempdoc = Path(tmp) / "processing.docx"
        shutil.copy2(src, tempdoc)
        try:
            pythoncom, app = _start_wps()
            doc = app.Documents.Open(str(tempdoc), ReadOnly=False, AddToRecentFiles=False)
            # Preflight all requested markers before changing the document.
            for row in manifest:
                _find_marker(doc, "[[MT:" + row["id"] + "]]")
            for row in manifest:
                outcome = _insert_one(app, doc, row, args.macro)
                report["equations"].append(outcome)
                report["converted"] += 1
                print(f"[{report['converted']}/{len(manifest)}] {row['id']} -> {outcome['ole_class']}", flush=True)
            leftovers = MARKER_RE.findall(str(doc.Content.Text))
            if leftovers:
                raise MathTypeError("文档还残留未处理的 MT 占位符：" + ", ".join(leftovers[:20]))
            doc.Save()
            doc.Close(SaveChanges=False)
            doc = None
            structure = inspect_docx(tempdoc)
            if structure["tex_translation_failed_previews"]:
                raise MathTypeError("发现疑似 TeX 翻译失败的公式预览：" +
                                    ", ".join(structure["tex_translation_failed_previews"]))
            # Do not publish the output before all equations have passed the OLE checks.
            if dst.exists() and args.overwrite:
                dst.unlink()
            shutil.move(str(tempdoc), str(dst))
            report["status"] = "COM_OLE_VERIFIED_VISUAL_REVIEW_REQUIRED"
            structure["file"] = str(dst)
            report["structural_check"] = structure
            return_code = 0
        except Exception as exc:
            report["status"] = "FAILED_NO_NEW_OUTPUT_PUBLISHED"
            report["error"] = str(exc)
            return_code = 2
            print(f"错误：{exc}", file=sys.stderr)
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=False)
                except Exception:
                    pass
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
            if pythoncom is not None:
                pythoncom.CoUninitialize()
            report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    _json_print(report)
    return return_code


def main():
    parser = argparse.ArgumentParser(description="WPS Writer + MathType OLE 自动化（拒绝非 OLE 降级）")
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("self-check", help="检查 Windows、pywin32、WPS COM、MathType OLE 注册")
    chk = sub.add_parser("inspect", help="辅助检查 docx 包内数学元素；非 OLE 类名最终验证")
    chk.add_argument("--input", required=True)
    conv = sub.add_parser("convert", help="按 JSON 公式清单处理现有 WPS 文档中的 [[MT:ID]]")
    conv.add_argument("--input", required=True, help="源 DOCX 路径（不会修改）")
    conv.add_argument("--output", required=True, help="全新的输出 DOCX 路径")
    conv.add_argument("--manifest", required=True, help="UTF-8 JSON 公式清单")
    conv.add_argument("--macro", default="MTCommand_TeXToggle", help="本机 WPS 已核实可调用的 MathType Toggle TeX 宏名")
    conv.add_argument("--overwrite", action="store_true", help="明确允许覆盖已有输出文件")
    ole = sub.add_parser("convert-ole", help="在同一个 WPS 工作副本中逐条转换 MathType OLE")
    ole.add_argument("--input", required=True, help="包含公式占位符的源 DOCX")
    ole.add_argument("--output", required=True, help="新的 DOCX 路径")
    ole.add_argument("--manifest", required=True, help="包含一条或多条公式的 UTF-8 JSON 清单")
    args = parser.parse_args()
    try:
        if args.command == "self-check":
            return self_check()
        if args.command == "inspect":
            _json_print(inspect_docx(args.input))
            return 0
        if args.command == "convert":
            return convert(args)
        if args.command == "convert-ole":
            from wps_ole_experiment import convert_ole
            return convert_ole(args, sys.modules[__name__])
    except (MathTypeError, OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"未预期错误：{exc}", file=sys.stderr)
        if os.environ.get("MATHTYPE_DEBUG") == "1":
            traceback.print_exc()
        return 2
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
