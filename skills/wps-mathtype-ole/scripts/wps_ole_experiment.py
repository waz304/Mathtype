"""Create MathType OLE objects in one WPS working copy."""

from __future__ import annotations

import ctypes
from ctypes import wintypes
from hashlib import sha256
import io
import json
import os
from pathlib import Path, PurePosixPath
import posixpath
import re
import shutil
import struct
import tempfile
import time
import zipfile
from xml.etree import ElementTree as ET

PROGID = "Equation.DSMT4"
EDITOR_CLASS = "EQNWINCLASS"
OLE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
OFFICE_NS = "urn:schemas-microsoft-com:office:office"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": WORD_NS, "m": "http://schemas.openxmlformats.org/officeDocument/2006/math"}
MATHML_NS = "http://www.w3.org/1998/Math/MathML"


class OleExperimentError(RuntimeError):
    pass


class EmptyEquation(OleExperimentError):
    pass


class UntranslatedInput(OleExperimentError):
    pass


class GuiThreadInfo(ctypes.Structure):
    _fields_ = [("size", wintypes.DWORD), ("flags", wintypes.DWORD)] + [
        (name, wintypes.HWND) for name in ("active", "focus", "capture", "menu", "move", "caret")
    ] + [("caret_rect", wintypes.RECT)]


class KeyInput(ctypes.Structure):
    _fields_ = [("vk", wintypes.WORD), ("scan", wintypes.WORD), ("flags", wintypes.DWORD),
                ("time", wintypes.DWORD), ("extra", ctypes.c_size_t)]


class MouseInput(ctypes.Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG), ("data", wintypes.DWORD),
                ("flags", wintypes.DWORD), ("time", wintypes.DWORD), ("extra", ctypes.c_size_t)]


class InputValue(ctypes.Union):
    _fields_ = [("key", KeyInput), ("mouse", MouseInput)]


class Input(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("value", InputValue)]


def _registration():
    import winreg

    def default_key(name):
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, name) as key:
            return winreg.QueryValueEx(key, "")[0]

    clsid = default_key(PROGID + r"\CLSID").strip("{}").upper()
    command = default_key(r"CLSID\{" + clsid + r"}\LocalServer32")
    executable = re.match(r'^\s*"?(.+?\.exe)(?:"|\s|$)', os.path.expandvars(command), re.I)
    if not executable:
        raise OleExperimentError(f"无法识别 {PROGID} 的 LocalServer32 可执行文件。")
    path = Path(executable.group(1)).resolve(strict=True)
    if path.name.lower() != "mathtype.exe":
        raise OleExperimentError(f"{PROGID} 的 LocalServer32 未指向 MathType.exe。")
    return clsid, path


class Desktop:
    def __init__(self, pythoncom, executable, name):
        import win32api
        import win32clipboard
        import win32con
        import win32gui
        import win32process

        self.com = pythoncom
        self.api = win32api
        self.cb = win32clipboard
        self.con = win32con
        self.gui = win32gui
        self.proc = win32process
        self.executable = executable
        self.name = name
        self.editor = None
        self.last_clipboard_sequence = None

    def wait(self, predicate, description, timeout=6):
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            self.com.PumpWaitingMessages()
            value = predicate()
            if value:
                return value
            time.sleep(0.05)
        raise OleExperimentError(f"等待 {description} 超时。")

    def _exe(self, hwnd):
        _, pid = self.proc.GetWindowThreadProcessId(hwnd)
        handle = self.api.OpenProcess(self.con.PROCESS_QUERY_INFORMATION | self.con.PROCESS_VM_READ, False, pid)
        try:
            return Path(self.proc.GetModuleFileNameEx(handle, 0)).resolve()
        finally:
            self.api.CloseHandle(handle)

    def _windows(self, cls):
        matches = []

        def collect(hwnd, _):
            if self.gui.IsWindowVisible(hwnd) and self.gui.GetClassName(hwnd) == cls:
                matches.append(hwnd)

        self.gui.EnumWindows(collect, None)
        return matches

    def wps_window(self):
        candidates = [
            hwnd for hwnd in self._windows("OpusApp")
            if self.name in self.gui.GetWindowText(hwnd) and self._exe(hwnd).name.lower() == "wps.exe"
        ]
        if len(candidates) > 1:
            raise OleExperimentError("找到多个同名测试 WPS 窗口。")
        return candidates[0] if candidates else None

    def editor_window(self):
        candidates = [
            hwnd for hwnd in self._windows(EDITOR_CLASS)
            if self.name in self.gui.GetWindowText(hwnd) and self._exe(hwnd) == self.executable
        ]
        if len(candidates) > 1:
            raise OleExperimentError("找到多个同名测试 MathType 编辑器。")
        return candidates[0] if candidates else None

    def focus(self):
        info = GuiThreadInfo()
        info.size = ctypes.sizeof(info)
        if (self.editor is None or self.gui.GetForegroundWindow() != self.editor
                or self.editor_window() != self.editor
                or not ctypes.windll.user32.GetGUIThreadInfo(0, ctypes.byref(info))
                or info.active != self.editor or info.focus != self.editor
                or info.flags & 0x1E or info.menu or info.capture):
            raise OleExperimentError("MathType 测试文档的公式画布未独占前台焦点，或菜单/对话框正在接管输入。")
        return {"hwnd": self.editor, "focus_hwnd": int(info.focus), "window_class": EDITOR_CLASS,
                "title": self.gui.GetWindowText(self.editor), "executable": str(self.executable)}

    def _send_input(self, events):
        batch = (Input * len(events))(*events)
        if ctypes.windll.user32.SendInput(len(events), batch, ctypes.sizeof(Input)) != len(events):
            raise OleExperimentError("Windows 未完整接收 MathType 前台输入。")

    def click_canvas(self):
        self.focus()
        vertical = []

        def collect(child, _):
            left, top, right, bottom = self.gui.GetWindowRect(child)
            if self.gui.GetClassName(child) == "ScrollBar" and bottom - top > right - left:
                vertical.append((left, top, right, bottom))

        self.gui.EnumChildWindows(self.editor, collect, None)
        if len(vertical) != 1:
            raise OleExperimentError("无法确定 MathType 公式画布滚动区。")
        left, top, _, bottom = vertical[0]
        client_left, _ = self.gui.ClientToScreen(self.editor, (0, 0))
        point = (client_left + min(20, (left - client_left) // 2), top + min(22, (bottom - top) // 2))
        if self.gui.WindowFromPoint(point) != self.editor:
            raise OleExperimentError("MathType 公式画布被遮挡。")
        if not ctypes.windll.user32.SetCursorPos(*point):
            raise OleExperimentError("无法将指针定位到 MathType 测试画布。")
        self.focus()
        self._send_input([Input(0, InputValue(mouse=MouseInput(flags=flag))) for flag in (2, 4)])
        self._settle_editor()
        return {"canvas_point_screen": point, **self.focus()}

    def _menu_command(self, group, command, label, key, wait_until_enabled=False):
        self.focus()
        menu = self.gui.GetSubMenu(self.gui.GetMenu(self.editor), group)
        for index in range(self.gui.GetMenuItemCount(menu)):
            if self.gui.GetMenuItemID(menu, index) != command:
                continue
            buffer = ctypes.create_unicode_buffer(512)
            ctypes.windll.user32.GetMenuStringW(
                wintypes.HMENU(menu), index, buffer, len(buffer), self.con.MF_BYPOSITION)
            if not re.search(label, buffer.value.replace("&", ""), re.I):
                break
            enabled = lambda: not self.gui.GetMenuState(menu, index, self.con.MF_BYPOSITION) & 3
            if wait_until_enabled:
                self.wait(enabled, buffer.value)
            elif not enabled():
                raise OleExperimentError(f"MathType 命令当前不可用：{buffer.value}")
            self.focus()
            ctrl = self.con.VK_CONTROL
            if any(self.api.GetAsyncKeyState(vk) & 0x8000 for vk in (ctrl, self.con.VK_SHIFT, self.con.VK_MENU)):
                raise OleExperimentError("检测到键盘修饰键已按下，停止前台输入。")
            self._send_input([Input(1, InputValue(key=KeyInput(vk=vk, flags=flags)))
                              for vk, flags in ((ctrl, 0), (key, 0), (key, 2), (ctrl, 2))])
            return buffer.value
        raise OleExperimentError(f"未找到经核对的 MathType 菜单命令 {command}。")

    def _clipboard_sequence(self):
        return self.cb.GetClipboardSequenceNumber()

    def _open_clipboard(self):
        from pywintypes import error as Win32Error

        def try_open():
            try:
                self.cb.OpenClipboard()
            except Win32Error as exc:
                if getattr(exc, "winerror", None) != 5:
                    raise
                return False
            return True

        self.wait(try_open, "剪贴板空闲", timeout=3)

    def _settle_editor(self):
        deadline = time.monotonic() + 0.35
        while time.monotonic() < deadline:
            self.com.PumpWaitingMessages()
            time.sleep(0.025)
        self.focus()

    def paste_and_read(self, tex, log):
        self.focus()
        cb = self.cb
        unicode_text = self.con.CF_UNICODETEXT
        self._open_clipboard()
        try:
            original = cb.GetClipboardData(unicode_text) if cb.IsClipboardFormatAvailable(unicode_text) else None
            formats = []
            fmt = cb.EnumClipboardFormats(0)
            while fmt:
                formats.append(fmt)
                fmt = cb.EnumClipboardFormats(fmt)
            cb.EmptyClipboard()
            cb.SetClipboardText(tex, unicode_text)
        finally:
            cb.CloseClipboard()
        self.last_clipboard_sequence = self._clipboard_sequence()
        try:
            self._open_clipboard()
            try:
                if cb.GetClipboardData(unicode_text) != tex:
                    raise OleExperimentError("粘贴前剪贴板内容发生变化。")
            finally:
                cb.CloseClipboard()
            log("paste", command=self._menu_command(1, 505, r"\bPaste\b|粘贴|貼上", ord("V")))
            self._settle_editor()
            self._menu_command(1, 513, r"\bSelect All\b|选择全部|全選", ord("A"))
            self._settle_editor()
            previous = self.last_clipboard_sequence
            self._menu_command(1, 504, r"\bCopy\b|复制|複製", ord("C"), wait_until_enabled=True)
            self.wait(lambda: self._clipboard_sequence() != previous, "MathType 复制结果")
            self._open_clipboard()
            try:
                owner = cb.GetClipboardOwner()
                if not owner or self.proc.GetWindowThreadProcessId(owner)[1] != self.proc.GetWindowThreadProcessId(self.editor)[1]:
                    raise OleExperimentError("剪贴板结果不是本次 MathType 编辑器产生的。")
                mtef = cb.GetClipboardData(cb.RegisterClipboardFormat("MathType EF"))
                mathml = cb.GetClipboardData(cb.RegisterClipboardFormat("MathML Presentation"))
            finally:
                cb.CloseClipboard()
            self.last_clipboard_sequence = self._clipboard_sequence()
            summary = _validate_mathml(mathml)
            if "".join(summary["mathml_tokens"].split()) == "".join(tex.split()):
                raise UntranslatedInput("MathType 把公式输入原样读回为文本。请检查 TeX 或提供有效的 MathML。")
            if not isinstance(mtef, bytes) or len(mtef) < 12 or mtef[0] != 5:
                raise OleExperimentError("MathType 剪贴板没有有效的 MTEF 5 数据。")
            log("translation", **summary, mtef_bytes=len(mtef), mtef_sha256=sha256(mtef).hexdigest())
            return mtef
        finally:
            # Restore ordinary text without clobbering a subsequent clipboard copy by the user.
            matches = self._clipboard_sequence() == self.last_clipboard_sequence
            self._open_clipboard()
            try:
                if matches and self._clipboard_sequence() == self.last_clipboard_sequence:
                    cb.EmptyClipboard()
                    if original is not None:
                        cb.SetClipboardText(original, unicode_text)
                    log("clipboard_cleanup", restored_text=original is not None,
                        replaced_non_text_formats=len(set(formats) - {1, 7, 13, 16}))
                else:
                    log("clipboard_cleanup", restored_text=False, external_change_detected=True)
            finally:
                cb.CloseClipboard()

    def close_editor(self):
        return self._menu_command(
            0, 402, r"(?:Close and Return|关闭并返回|關閉並返回).*" + re.escape(self.name), self.con.VK_F4)


def _validate_mathml(blob):
    root = ET.fromstring(blob.rstrip(b"\0"))
    if root.tag != f"{{{MATHML_NS}}}math":
        raise OleExperimentError("MathType 剪贴板的 MathML 缺少数学根元素。")
    elements = []
    tokens = []
    for node in root.iter():
        name = node.tag.rsplit("}", 1)[-1]
        if name in {"annotation", "annotation-xml"}:
            continue
        color = re.sub(r"\s", "", node.get("mathcolor", node.get("color", ""))).lower()
        if name == "merror" or color in {"red", "#ff0000", "#f00", "rgb(255,0,0)"}:
            raise OleExperimentError("MathType 翻译结果包含错误标记或红色诊断。")
        if name in {"mi", "mo", "mn", "mtext"} and node.text:
            tokens.append(node.text)
        elements.append(name)
    text = "".join(tokens)
    if not text.strip():
        raise EmptyEquation("MathType 编辑器回读了空公式。")
    if re.search(r"translation failed|\\[a-z]+", text, re.I):
        raise OleExperimentError("MathType 翻译结果仍包含错误文本或未解析的 TeX 命令。")
    return {"mathml_tokens": text, "mathml_elements": sorted(set(elements))}


def _verify_shape(shape):
    cls = str(shape.OLEFormat.ClassType)
    progid = str(shape.OLEFormat.ProgID)
    if int(shape.Type) != 1 or {cls.lower(), progid.lower()} != {PROGID.lower()}:
        raise OleExperimentError(f"WPS 对象不是预期内嵌类型：Type={shape.Type}, ClassType={cls}, ProgID={progid}")
    return {"class_type": cls, "progid": progid, "shape_type": int(shape.Type)}


def _document_body(root):
    return root.find("w:body", NS)


def _paragraph_content(paragraph):
    parts = []
    for run in paragraph.findall("w:r", NS):
        for child in run:
            if child.tag == f"{{{WORD_NS}}}t":
                parts.append(child.text or "")
            elif child.tag == f"{{{WORD_NS}}}tab":
                parts.append("\t")
            elif child.tag == f"{{{WORD_NS}}}object":
                parts.append("\ufffc")
    return "".join(parts)


def _plan_batch(source, rows):
    """Locate each marker by paragraph and object index, including inline siblings."""
    with zipfile.ZipFile(source) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    wanted = {"[[MT:" + row["id"] + "]]": row for row in rows}
    pattern = re.compile("|".join(map(re.escape, wanted)))
    paragraphs = []
    locations = {}
    for index, paragraph in enumerate(_document_body(root).findall("w:p", NS)):
        text = _paragraph_content(paragraph)
        expected = []
        offset = 0
        object_index = 0
        for match in pattern.finditer(text):
            row = wanted[match.group()]
            ident = row["id"]
            if ident in locations:
                raise OleExperimentError(f"{ident}: 主体中有重复占位符。")
            if row["kind"] == "display" and text != match.group():
                raise OleExperimentError(f"{ident}: 独立公式占位符必须独占一个段落。")
            prefix = text[offset:match.start()]
            expected.append(prefix)
            object_index += prefix.count("\ufffc")
            locations[ident] = (index, object_index)
            object_index += 1
            if row["kind"] == "display":
                expected.append("\t\ufffc" + ("\t" + row["number"] if row["number"] else ""))
            else:
                expected.append("\ufffc")
            offset = match.end()
        paragraphs.append("".join(expected) + text[offset:])
    missing = [row["id"] for row in rows if row["id"] not in locations]
    if missing:
        raise OleExperimentError("未在主体普通段落中找到占位符：" + "、".join(missing))
    return paragraphs, locations, len(root.findall(".//w:object", NS)), len(root.findall(".//m:oMath", NS))


def _verify_embedding(package, element, relations, clsid, mtef):
    import olefile

    if (element is None or element.get("ProgID", "").lower() != PROGID.lower()
            or element.get("Type") != "Embed"):
        raise OleExperimentError("占位符对应位置缺少内嵌 MathType 对象。")
    rid = element.get(f"{{{OLE_REL}}}id")
    link = relations.get(rid)
    if link is None or link.get("Type") != OLE_REL + "/oleObject" or link.get("TargetMode") == "External":
        raise OleExperimentError("内嵌 MathType 对象缺少有效的 DOCX 关系。")
    target = posixpath.normpath(posixpath.join("word", link.get("Target", "")))
    if PurePosixPath(target).parent != PurePosixPath("word/embeddings"):
        raise OleExperimentError("MathType 对象目标不在 DOCX 内嵌目录中。")
    with olefile.OleFileIO(io.BytesIO(package.read(target))) as storage:
        if storage.root.clsid.upper() != clsid or not storage.exists("Equation Native"):
            raise OleExperimentError("DOCX 内嵌对象缺少正确的 CLSID 或 Equation Native 流。")
        native = storage.openstream("Equation Native").read()
    if len(native) < 28:
        raise OleExperimentError("DOCX 的 Equation Native 流不完整。")
    header_size, = struct.unpack_from("<H", native)
    data_size, = struct.unpack_from("<I", native, 8)
    if header_size != 28 or native[header_size:header_size + data_size] != mtef:
        raise OleExperimentError("DOCX 的 MTEF 内容与 MathType 编辑器回读结果不一致。")
    return {"progid": element.get("ProgID"), "relationship_id": rid, "embedding": target,
            "storage_clsid": clsid, "stream": "Equation Native", "mtef_matches_editor": True}


def _verify_package(saved, plan, converted, clsid, log):
    expected, locations, previous_ole, previous_omml = plan
    with zipfile.ZipFile(saved) as package:
        root = ET.fromstring(package.read("word/document.xml"))
        if len(root.findall(".//w:object", NS)) != previous_ole + len(converted):
            raise OleExperimentError("DOCX 主体的新增内嵌对象数与本批公式数不符。")
        if len(root.findall(".//m:oMath", NS)) != previous_omml:
            raise OleExperimentError("DOCX 的 OMML 公式个数发生变化。")
        paragraphs = _document_body(root).findall("w:p", NS)
        if [_paragraph_content(para) for para in paragraphs] != expected:
            raise OleExperimentError("保存后公式位置、编号或正文与预期不符。")
        relations = {rel.get("Id"): rel for rel in ET.fromstring(package.read("word/_rels/document.xml.rels"))}
        for ident, (mtef, expected_tabs) in converted.items():
            index, object_index = locations[ident]
            paragraph = paragraphs[index]
            element = paragraph.findall("w:r/w:object", NS)[object_index].find(f"{{{OFFICE_NS}}}OLEObject")
            tab_positions = {}
            if expected_tabs is not None:
                for tab in paragraph.findall("w:pPr/w:tabs/w:tab", NS):
                    tab_positions[tab.get(f"{{{WORD_NS}}}val")] = int(tab.get(f"{{{WORD_NS}}}pos"))
                if any(abs(tab_positions.get(kind, -1000) - pos) > 2 for kind, pos in expected_tabs.items()):
                    raise OleExperimentError(f"{ident}: 保存后的制表位不符：{tab_positions}")
            try:
                details = _verify_embedding(package, element, relations, clsid, mtef)
            except Exception as exc:
                raise OleExperimentError(f"{ident}: {exc}") from exc
            log("verify_docx", formula_id=ident, paragraph=index + 1,
                tab_stops_twips=tab_positions, omml_added=0, **details)


def _layout(doc, marker, position, row, legacy):
    found = legacy._find_marker(doc, marker)
    if str(found.Text) != marker or int(found.Start) <= position:
        raise OleExperimentError("临时插入后，原占位符的位置或内容不符。")
    found.Text = "\t" + row["number"] if row["kind"] == "display" and row["number"] else ""
    if row["kind"] == "inline":
        return None
    doc.Range(position, position).InsertBefore("\t")
    legacy._format_display_paragraph(doc, position)
    para = doc.Range(position, position).Paragraphs.Item(1)
    return {("center" if int(tab.Alignment) == 1 else "right"): round(float(tab.Position) * 20)
            for tab in (para.Format.TabStops.Item(1), para.Format.TabStops.Item(2))}


def _convert_one(app, doc, wps, desktop, row, legacy, log):
    marker = "[[MT:" + row["id"] + "]]"
    found = legacy._find_marker(doc, marker)
    if str(found.Text) != marker:
        raise OleExperimentError("WPS Find 返回的范围与占位符不符。")
    if row["kind"] == "display" and str(found.Paragraphs.Item(1).Range.Text).rstrip("\r") != marker:
        raise OleExperimentError("独立公式占位符必须独占一个测试段落。")
    position = int(found.Start)
    target = doc.Range(position, position)
    # WPS can use Selection even when Range is provided to AddOLEObject.
    target.Select()
    if int(app.Selection.Start) != position or int(app.Selection.End) != position:
        raise OleExperimentError("WPS 选区与待插入位置不符。")
    before = doc.InlineShapes.Count
    shape = doc.InlineShapes.AddOLEObject(ClassType=PROGID, DisplayAsIcon=False, Range=target)
    descriptors = _verify_shape(shape)
    if doc.InlineShapes.Count != before + 1 or int(shape.Range.Start) != position:
        raise OleExperimentError(f"WPS 未在测试占位符前插入一条 OLE，位置 {shape.Range.Start}，预期 {position}。")
    log("add_ole", position=position, **descriptors)
    desktop.editor = desktop.editor_window()
    if desktop.editor is None:
        shape.OLEFormat.DoVerb(-2)  # OLEIVERB_OPEN.
        desktop.editor = desktop.wait(desktop.editor_window, "MathType OLE 编辑器")
    if desktop.gui.GetForegroundWindow() != desktop.editor and desktop.gui.IsWindowEnabled(desktop.editor):
        from pywintypes import error as Win32Error

        try:
            desktop.gui.SetForegroundWindow(desktop.editor)
        except Win32Error:
            pass  # Windows can decline activation. The focus check below still applies.
    desktop.wait(lambda: desktop.gui.GetForegroundWindow() == desktop.editor, "MathType 编辑器进入前台")
    log("focus", **desktop.click_canvas())
    clipboard_text = row["mathml"] if "mathml" in row else "$" + row["latex"] + "$"
    try:
        mtef = desktop.paste_and_read(clipboard_text, log)
    except (EmptyEquation, UntranslatedInput) as exc:
        # paste_and_read selected the full editor contents to copy the result.
        # Select again before replacing the failed paste in the same new object.
        log("retry_paste", reason=str(exc))
        desktop._menu_command(1, 513, r"\bSelect All\b|选择全部|全選", ord("A"))
        desktop._settle_editor()
        mtef = desktop.paste_and_read(clipboard_text, log)
    command = desktop.close_editor()
    desktop.wait(lambda: desktop.editor_window() is None, "MathType 关闭返回")
    desktop.editor = None
    desktop.wait(lambda: desktop.gui.GetForegroundWindow() == wps, "WPS 测试文档返回前台")
    log("return_to_wps", command=command, hwnd=wps, **_verify_shape(shape))
    tabs = _layout(doc, marker, position, row, legacy)
    log("layout", kind=row["kind"], number=row["number"], expected_tab_stops_twips=tabs)
    return mtef, tabs


def _run(src, dst, rows, directory, legacy, log):
    plan = _plan_batch(src, rows)
    clsid, executable = _registration()
    log("registration", progid=PROGID, clsid=clsid, server=str(executable))
    import olefile  # Check the verification dependency before opening an application.
    pythoncom, client = legacy._win32_client()
    tmp = directory / (directory.name + ".docx")
    shutil.copyfile(src, tmp)
    desktop = Desktop(pythoncom, executable, tmp.name)
    if ctypes.windll.ole32.OleInitialize(None) not in (0, 1):
        raise OleExperimentError("无法初始化当前桌面线程的 OLE。")
    doc = None
    owned_wps = False
    try:
        existing_pids = set(desktop.proc.EnumProcesses())
        app = client.DispatchEx("kwps.application")
        if app.Documents.Count:
            raise OleExperimentError("WPS COM 会话已有打开的文档。请保留当前工作并关闭该会话后再试。")
        app.Visible = True
        doc = app.Documents.Open(str(tmp), ReadOnly=False, AddToRecentFiles=False)
        if Path(doc.FullName).resolve() != tmp or doc.ReadOnly:
            raise OleExperimentError("WPS 未以可写方式打开指定测试副本。")
        doc.Activate()
        wps = desktop.wait(desktop.wps_window, "测试副本 WPS 窗口")
        _, pid = desktop.proc.GetWindowThreadProcessId(wps)
        owned_wps = pid not in existing_pids
        log("wps", com_version=str(app.Version), com_app_name=str(app.Name),
            hwnd=wps, pid=pid, owned_process=owned_wps,
            executable=str(desktop._exe(wps)), olefile_version=olefile.__version__)
        # Validate WPS can resolve all targets before editing any of them.
        for row in rows:
            legacy._find_marker(doc, "[[MT:" + row["id"] + "]]")
        converted = {}
        for row in rows:
            ident = row["id"]
            def formula_log(step, **details):
                log(step, formula_id=ident, **details)
            try:
                converted[ident] = _convert_one(app, doc, wps, desktop, row, legacy, formula_log)
            except Exception as exc:
                raise OleExperimentError(f"{ident}: {exc}") from exc
        doc.Save()
        doc.Close(SaveChanges=False)
        doc = None
        log("save", working_copy=str(tmp), formulas=len(converted))
        _verify_package(tmp, plan, converted, clsid, log)
        os.rename(tmp, dst)
        log("publish", file=str(dst), manual_double_click="required")
    finally:
        try:
            if doc is not None and desktop.editor is not None and desktop.editor_window() == desktop.editor:
                desktop.close_editor()
                desktop.wait(lambda: desktop.editor_window() is None, "测试编辑器清理")
        except Exception as exc:
            log("cleanup_editor", status="failed", error=str(exc))
        try:
            if doc is not None:
                doc.Close(SaveChanges=False)
            if owned_wps and not app.Documents.Count:
                app.Quit()
        except Exception as exc:
            log("cleanup_wps", status="failed", error=str(exc))
        finally:
            ctypes.windll.ole32.OleUninitialize()


def convert_ole(args, legacy):
    src = Path(args.input).expanduser().resolve()
    dst = Path(args.output).expanduser().resolve()
    if src == dst or dst.suffix.lower() != ".docx":
        raise OleExperimentError("输出须是与测试源文件不同的 .docx 路径。")
    dst.parent.mkdir(parents=True, exist_ok=True)
    report_path = dst.with_name(dst.stem + "_ole_report.json")
    report = {"mode": "convert-ole", "source": str(src), "output": str(dst), "converted": 0,
              "status": "FAILED_NO_NEW_OUTPUT_PUBLISHED", "steps": []}

    def log(step, **details):
        entry = {"step": step, **details}
        report["steps"].append(entry)
        print(json.dumps(entry, ensure_ascii=False), flush=True)

    source_hash = None
    try:
        if not src.is_file() or src.suffix.lower() != ".docx" or dst.exists():
            raise OleExperimentError("需要存在的测试 DOCX 和尚未存在的输出路径。")
        source_hash = sha256(src.read_bytes()).hexdigest()
        manifest = legacy._load_manifest(args.manifest, allow_mathml=True)
        for row in manifest:
            if "mathml" in row:
                try:
                    _validate_mathml(row["mathml"].encode("utf-8"))
                except (ET.ParseError, OleExperimentError) as exc:
                    raise OleExperimentError(f"{row['id']}: 无效的 MathML：{exc}") from exc
        report["requested"] = len(manifest)
        log("preflight", formula_ids=[row["id"] for row in manifest], source_sha256=source_hash)
        with tempfile.TemporaryDirectory(prefix="mt_ole_", dir=dst.parent, ignore_cleanup_errors=True) as directory:
            _run(src, dst, manifest, Path(directory), legacy, log)
        report.update(status="COM_OLE_VERIFIED_MANUAL_DOUBLE_CLICK_REQUIRED", converted=len(manifest))
    except Exception as exc:
        report["error"] = str(exc)
        log("error", type=type(exc).__name__, message=str(exc))
    finally:
        report["source_unchanged"] = source_hash is not None and src.is_file() and sha256(src.read_bytes()).hexdigest() == source_hash
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps({"status": report["status"], "converted": report["converted"], "report": str(report_path)},
                         ensure_ascii=False), flush=True)
    return 0 if report["converted"] else 2
