"""ファイルを開く (可能なら該当位置へジャンプする)。

- テキスト      : 設定したエディタで行・桁を指定して開く (未設定なら既定のアプリ)
- Excel        : pywin32 があればシート・セルを選択した状態で開く
- Word         : pywin32 があれば該当段落を選択した状態で開く
- PowerPoint   : pywin32 があれば該当スライドを表示して開く
- PDF / その他 : 既定のアプリで開く
"""
from __future__ import annotations

import os
import re
import shlex
import subprocess
import sys
import threading

IS_WIN = sys.platform.startswith("win")

EDITOR_PRESETS = [
    ("（既定のアプリで開く）", ""),
    ("サクラエディタ", r'"C:\Program Files (x86)\sakura\sakura.exe" -Y={line} -X={col} "{path}"'),
    ("秀丸エディタ", r'"C:\Program Files\Hidemaru\Hidemaru.exe" /j{line},{col} "{path}"'),
    ("Visual Studio Code", r'code -g "{path}:{line}:{col}"'),
    ("Notepad++", r'"C:\Program Files\Notepad++\notepad++.exe" -n{line} -c{col} "{path}"'),
]


def detect_editor() -> str:
    """よく使われるエディタがインストールされていれば、そのコマンドを返す"""
    if not IS_WIN:
        return ""
    for name, cmd in EDITOR_PRESETS[1:]:
        m = re.match(r'"([^"]+)"', cmd)
        if m and os.path.exists(m.group(1)):
            return cmd
    return ""


def open_default(path: str):
    if IS_WIN:
        os.startfile(path)  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def reveal_in_explorer(path: str):
    if IS_WIN:
        subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
    else:
        open_default(os.path.dirname(path))


def _run_editor(template: str, path: str, line, col):
    cmd = template.format(path=path, line=line or 1, col=col or 1)
    if IS_WIN:
        subprocess.Popen(cmd, shell=False)
    else:
        subprocess.Popen(shlex.split(cmd))


def _sheet_name(place: str) -> str:
    return re.sub(r"\(非表示\)$", "", place)


_SLIDE_RE = re.compile(r"スライド(\d+)")
_CELL_RE = re.compile(r"^([A-Z]+\d+)")


def _com_excel(path, place, label):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = True
        wb = None
        for w in xl.Workbooks:
            if os.path.normcase(w.FullName) == os.path.normcase(path):
                wb = w
                break
        if wb is None:
            wb = xl.Workbooks.Open(path)
        m = _CELL_RE.match(label or "")
        try:
            ws = wb.Worksheets(_sheet_name(place))
            ws.Activate()
            if m:
                ws.Range(m.group(1)).Select()
                xl.ActiveWindow.ScrollRow = max(1, ws.Range(m.group(1)).Row - 5)
        except Exception:
            pass  # 非表示シート等は選択できないのでブックを開くだけ
        _activate(xl)
    finally:
        pythoncom.CoUninitialize()


def _com_word(path, line, col, match_len):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        wd = win32com.client.Dispatch("Word.Application")
        wd.Visible = True
        doc = None
        for d in wd.Documents:
            if os.path.normcase(d.FullName) == os.path.normcase(path):
                doc = d
                break
        if doc is None:
            doc = wd.Documents.Open(path)
        try:
            if line and line <= doc.Paragraphs.Count:
                rng = doc.Paragraphs(line).Range
                start = rng.Start + max(0, (col or 1) - 1)
                doc.Range(start, min(rng.End, start + max(1, match_len))).Select()
        except Exception:
            pass
        doc.Activate()
        _activate(wd)
    finally:
        pythoncom.CoUninitialize()


def _com_ppt(path, place):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    try:
        pp = win32com.client.Dispatch("PowerPoint.Application")
        pres = None
        for p in pp.Presentations:
            if os.path.normcase(p.FullName) == os.path.normcase(path):
                pres = p
                break
        if pres is None:
            pres = pp.Presentations.Open(path)
        m = _SLIDE_RE.search(place or "")
        if m:
            try:
                pres.Windows(1).View.GotoSlide(int(m.group(1)))
            except Exception:
                pass
        pp.Activate()
    finally:
        pythoncom.CoUninitialize()


def _activate(app):
    try:
        import win32gui
        hwnd = app.Hwnd if hasattr(app, "Hwnd") else None
        if hwnd:
            win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass


def _has_pywin32() -> bool:
    if not IS_WIN:
        return False
    try:
        import win32com.client  # noqa: F401
        return True
    except Exception:
        return False


def open_at(path: str, kind: str, place: str = "", line=None, col=None, label: str = "",
            match: str = "", editor: str = "", on_error=None):
    """該当箇所を開く。位置指定ができない場合は普通にファイルを開く。"""
    ext = os.path.splitext(path)[1].lower()

    def fallback():
        try:
            open_default(path)
        except Exception as e:
            if on_error:
                on_error(f"ファイルを開けませんでした: {e}")

    if kind == "text":
        if editor:
            try:
                _run_editor(editor, path, line, col)
                return
            except Exception as e:
                if on_error:
                    on_error(f"エディタを起動できませんでした ({e})。既定のアプリで開きます。")
        fallback()
        return

    if kind in ("excel", "word", "ppt") and _has_pywin32() and ext not in (".csv",):
        def job():
            try:
                if kind == "excel":
                    _com_excel(path, place, label)
                elif kind == "word":
                    _com_word(path, line if place == "本文" or place.startswith("表") else None,
                              col, len(match or ""))
                else:
                    _com_ppt(path, place)
            except Exception:
                fallback()
        threading.Thread(target=job, daemon=True).start()
        return
    fallback()
