"""各ファイル形式からテキストを「位置情報付きの単位(Unit)」として取り出す。

Unit は次のタプル:
    (mode, place, line, col, label, text)

    mode  'L' : 複数行テキスト。line は先頭行番号。ヒット位置の行/桁は text 内の改行から計算する
                (テキストファイル、PDFのページ、PowerPointのスライドなど)
          'F' : 位置固定。line/col はそのまま使う (Excelのセル、Wordの段落など)
                col が None の場合は「単位内の文字位置」を列として使う
    place : シート名 / "3ページ" / "スライド2" など (テキストファイルは "")
    label : 表示用の位置ラベル ("B12", "表1 (2行,3列)", "図形: 四角形 1" など)

このモジュールは Qt に依存しない (検索ワーカープロセスで import されるため)。
"""
from __future__ import annotations

import codecs
import os
import re
import zipfile
import posixpath
import xml.etree.ElementTree as ET

# ---------------------------------------------------------------- 種別判定
KIND_TEXT, KIND_EXCEL, KIND_WORD, KIND_PPT, KIND_PDF = "text", "excel", "word", "ppt", "pdf"
KIND_LABEL = {
    KIND_TEXT: "テキスト", KIND_EXCEL: "Excel", KIND_WORD: "Word",
    KIND_PPT: "PowerPoint", KIND_PDF: "PDF",
}

EXCEL_NEW = {".xlsx", ".xlsm", ".xltx", ".xltm"}
EXCEL_OLD = {".xls", ".xlt"}
WORD_NEW = {".docx", ".docm", ".dotx", ".dotm"}
WORD_OLD = {".doc", ".dot"}
PPT_NEW = {".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm"}
PPT_OLD = {".ppt", ".pot", ".pps"}
PDF_EXT = {".pdf"}

# 中身を見るまでもなくバイナリと分かる拡張子 (テキストとして読まない)
BINARY_EXT = {
    ".exe", ".dll", ".so", ".lib", ".obj", ".o", ".a", ".pdb", ".class", ".jar", ".war",
    ".zip", ".7z", ".rar", ".gz", ".tgz", ".bz2", ".xz", ".lzh", ".cab", ".iso", ".msi",
    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".ico", ".tif", ".tiff", ".webp", ".psd", ".emf", ".wmf",
    ".mp3", ".mp4", ".wav", ".avi", ".mov", ".wmv", ".flv", ".mkv", ".m4a",
    ".ttf", ".otf", ".woff", ".woff2", ".pyc", ".pyd", ".db", ".sqlite", ".mdb", ".accdb",
    ".xlsb", ".vsd", ".vsdx", ".one", ".msg", ".pst", ".ost", ".bin", ".dat", ".lnk",
}


def kind_of(path: str) -> str | None:
    ext = os.path.splitext(path)[1].lower()
    if ext in EXCEL_NEW or ext in EXCEL_OLD:
        return KIND_EXCEL
    if ext in WORD_NEW or ext in WORD_OLD:
        return KIND_WORD
    if ext in PPT_NEW or ext in PPT_OLD:
        return KIND_PPT
    if ext in PDF_EXT:
        return KIND_PDF
    if ext in BINARY_EXT:
        return None
    return KIND_TEXT


class ExtractError(Exception):
    """利用者に見せる理由付きの読み取りエラー"""


# ---------------------------------------------------------------- テキスト
ENCODING_CHOICES = [
    ("自動選択", "auto"), ("UTF-8", "utf-8"), ("Shift_JIS (CP932)", "cp932"),
    ("EUC-JP", "euc_jp"), ("UTF-16LE", "utf-16-le"), ("UTF-16BE", "utf-16-be"),
    ("JIS (ISO-2022-JP)", "iso2022_jp"),
]
ENCODING_DISPLAY = {
    "utf-8": "UTF-8", "utf-8-sig": "UTF-8(BOM)", "cp932": "SJIS", "euc_jp": "EUC-JP",
    "utf-16-le": "UTF-16LE", "utf-16-be": "UTF-16BE", "utf-16": "UTF-16",
    "iso2022_jp": "JIS", "ascii": "ASCII",
}


def decode_text(data: bytes, encoding: str = "auto") -> tuple[str, str]:
    """バイト列をデコードして (テキスト, 文字コード名) を返す。バイナリなら ExtractError。"""
    if encoding != "auto":
        if data.startswith(codecs.BOM_UTF8) and encoding == "utf-8":
            return data[3:].decode("utf-8", "replace"), "utf-8-sig"
        return data.decode(encoding, "replace"), encoding
    if data.startswith(codecs.BOM_UTF8):
        return data[3:].decode("utf-8", "replace"), "utf-8-sig"
    if data.startswith(codecs.BOM_UTF16_LE):
        return data[2:].decode("utf-16-le", "replace"), "utf-16-le"
    if data.startswith(codecs.BOM_UTF16_BE):
        return data[2:].decode("utf-16-be", "replace"), "utf-16-be"
    head = data[:8192]
    if b"\x00" in head:
        # BOM 無し UTF-16 の簡易判定 (ASCII 主体なら偶数/奇数位置に NUL が偏る)
        if len(head) >= 64:
            even = head[0::2].count(0)
            odd = head[1::2].count(0)
            n = len(head) // 2
            if odd > n * 0.4 and even < n * 0.05:
                return data.decode("utf-16-le", "replace"), "utf-16-le"
            if even > n * 0.4 and odd < n * 0.05:
                return data.decode("utf-16-be", "replace"), "utf-16-be"
        raise ExtractError("バイナリファイルのためスキップ")
    try:
        text = data.decode("utf-8")
        return text, ("ascii" if text.isascii() else "utf-8")
    except UnicodeDecodeError:
        pass
    if b"\x1b$" in head:
        try:
            return data.decode("iso2022_jp"), "iso2022_jp"
        except UnicodeDecodeError:
            pass
    # SJIS と EUC-JP は互いに「読めてしまう」ことが多いので、両方試して日本語らしい方を選ぶ
    cands = []
    for enc in ("cp932", "euc_jp"):
        try:
            cands.append((_ja_score(data.decode(enc)), enc))
        except UnicodeDecodeError:
            pass
    if cands:
        cands.sort(key=lambda x: -x[0])
        enc = cands[0][1]
        return data.decode(enc), enc
    # どれでも厳密に読めない場合は CP932 で読めるところだけ読む
    return data.decode("cp932", "replace"), "cp932"


def _ja_score(text: str) -> int:
    """ひらがな・カタカナ・常用漢字域は加点、半角カナ・外字域・まれな記号は減点"""
    score = 0
    for ch in text[:20000]:
        o = ord(ch)
        if o < 0x80:
            continue
        if 0x3040 <= o <= 0x30FF or 0x4E00 <= o <= 0x9FFF or 0x3000 <= o <= 0x303F or 0xFF01 <= o <= 0xFF5E:
            score += 2
        elif 0xFF61 <= o <= 0xFF9F:
            score -= 1
        else:
            score -= 3
    return score


def read_text_file(path: str, encoding: str = "auto") -> tuple[str, str]:
    with open(path, "rb") as f:
        data = f.read()
    return decode_text(data, encoding)


# ---------------------------------------------------------------- 共通 (OOXML)
def _open_zip(path: str) -> zipfile.ZipFile:
    try:
        return zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        # 暗号化された OOXML は OLE 形式 (EncryptionInfo ストリームを持つ)
        try:
            import olefile
            if olefile.isOleFile(path):
                ole = olefile.OleFileIO(path)
                try:
                    if ole.exists("EncryptionInfo"):
                        raise ExtractError("パスワード保護されているため読めません")
                finally:
                    ole.close()
        except ExtractError:
            raise
        except Exception:
            pass
        raise ExtractError("ファイルが壊れているか、形式が拡張子と一致しません")


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _rels(z: zipfile.ZipFile, part: str) -> dict[str, str]:
    """part (例 'xl/workbook.xml') の .rels を読み、rId -> 絶対パート名 の辞書を返す"""
    d, name = posixpath.split(part)
    rel_path = posixpath.join(d, "_rels", name + ".rels")
    out: dict[str, str] = {}
    try:
        root = ET.fromstring(z.read(rel_path))
    except KeyError:
        return out
    for r in root:
        if r.get("TargetMode") == "External":
            continue
        tgt = r.get("Target", "")
        if tgt.startswith("/"):
            full = tgt.lstrip("/")
        else:
            full = posixpath.normpath(posixpath.join(d, tgt))
        out[r.get("Id")] = full
        out.setdefault("type:" + r.get("Type", "").rsplit("/", 1)[-1], full)
    return out


def _rels_by_type(z: zipfile.ZipFile, part: str, typ: str) -> list[str]:
    d, name = posixpath.split(part)
    rel_path = posixpath.join(d, "_rels", name + ".rels")
    try:
        root = ET.fromstring(z.read(rel_path))
    except KeyError:
        return []
    res = []
    for r in root:
        if r.get("Type", "").endswith("/" + typ) and r.get("TargetMode") != "External":
            tgt = r.get("Target", "")
            res.append(tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join(d, tgt)))
    return res


XDR_NS = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"


def _xlsx_drawing_units(xml: bytes, place: str) -> list[tuple]:
    """Excel の図形 (テキストボックス・吹き出し・グループ内の図形など) の文字を、置かれたセル位置付きで返す"""
    units = []
    root = ET.fromstring(xml)
    for anc in root:
        if _local(anc.tag) not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
            continue
        where = ""
        fr = anc.find(XDR_NS + "from")
        if fr is not None:
            try:
                c = int(fr.findtext(XDR_NS + "col")) + 1
                r = int(fr.findtext(XDR_NS + "row")) + 1
                where = f"（{col_letter(c)}{r}付近）"
            except (TypeError, ValueError):
                pass
        for sp in anc.iter():
            if _local(sp.tag) not in ("sp", "graphicFrame"):
                continue
            name = ""
            for c in sp.iter():
                if _local(c.tag) == "cNvPr":
                    name = c.get("name", "")
                    break
            for p in sp.iter(A_NS_ + "p"):
                t = _a_paragraph_text(p)
                if t.strip():
                    units.append(("F", place, None, None, f"図形: {name or '(名前なし)'}{where}", t))
    return units


A_NS_ = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _a_paragraph_text(p) -> str:
    parts = []
    for n in p.iter():
        ln = _local(n.tag)
        if ln == "t" and n.text:
            parts.append(n.text)
        elif ln == "br":
            parts.append(" ")
    return "".join(parts)


# ---------------------------------------------------------------- Excel
def col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


_CELL_RE = re.compile(r"([A-Z]+)(\d+)")


def _col_number(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def extract_xlsx(path: str) -> list[tuple]:
    z = _open_zip(path)
    units: list[tuple] = []
    with z:
        names = set(z.namelist())
        # 共有文字列 (ふりがな rPh は除外)
        shared: list[str] = []
        ss_part = "xl/sharedStrings.xml"
        wb_rels = _rels(z, "xl/workbook.xml")
        for k, v in wb_rels.items():
            if k == "type:sharedStrings":
                ss_part = v
        if ss_part in names:
            with z.open(ss_part) as f:
                for ev, el in ET.iterparse(f, events=("end",)):
                    if _local(el.tag) == "si":
                        buf = []
                        for ch in el:
                            cl = _local(ch.tag)
                            if cl == "t":
                                buf.append(ch.text or "")
                            elif cl == "r":
                                for t in ch:
                                    if _local(t.tag) == "t":
                                        buf.append(t.text or "")
                        shared.append("".join(buf))
                        el.clear()
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheets = []
        for el in wb.iter():
            if _local(el.tag) == "sheet":
                rid = None
                for k, v in el.attrib.items():
                    if k.endswith("}id"):
                        rid = v
                sheets.append((el.get("name", "?"), el.get("state", "visible"), wb_rels.get(rid)))
        for sname, state, part in sheets:
            if not part or part not in names:
                continue
            place = sname if state == "visible" else f"{sname}(非表示)"
            row_no = 0
            with z.open(part) as f:
                col_no = 0
                for ev, el in ET.iterparse(f, events=("end",)):
                    ln = _local(el.tag)
                    if ln == "c":
                        ref = el.get("r")
                        if ref:
                            m = _CELL_RE.match(ref)
                            col_no, row_no = _col_number(m.group(1)), int(m.group(2))
                        else:
                            col_no += 1
                        t = el.get("t", "n")
                        val = None
                        for ch in el:
                            cl = _local(ch.tag)
                            if cl == "v":
                                val = ch.text
                            elif cl == "is":
                                val = "".join((x.text or "") for x in ch.iter() if _local(x.tag) == "t"
                                              and not _in_rph(ch, x))
                        if val is not None and val != "":
                            if t == "s":
                                try:
                                    val = shared[int(val)]
                                except (ValueError, IndexError):
                                    pass
                            elif t == "b":
                                val = "TRUE" if val == "1" else "FALSE"
                            if val:
                                units.append(("F", place, row_no, col_no, f"{col_letter(col_no)}{row_no}", val))
                        el.clear()
                    elif ln == "row":
                        r = el.get("r")
                        if r:
                            row_no = int(r)
                        col_no = 0
                        el.clear()
            # コメント
            for cpart in _rels_by_type(z, part, "comments"):
                if cpart in names:
                    root = ET.fromstring(z.read(cpart))
                    for cm in root.iter():
                        if _local(cm.tag) == "comment":
                            ref = cm.get("ref", "")
                            txt = "".join((x.text or "") for x in cm.iter() if _local(x.tag) == "t")
                            m = _CELL_RE.match(ref)
                            if txt and m:
                                units.append(("F", place, int(m.group(2)), _col_number(m.group(1)),
                                              f"{ref} (コメント)", txt))
            # 図形・テキストボックス
            for dpart in _rels_by_type(z, part, "drawing"):
                if dpart in names:
                    units.extend(_xlsx_drawing_units(z.read(dpart), place))
    return units


def _in_rph(parent, node) -> bool:
    for x in parent.iter():
        if _local(x.tag) == "rPh" and any(y is node for y in x.iter()):
            return True
    return False


def extract_xls(path: str) -> list[tuple]:
    import xlrd
    try:
        book = xlrd.open_workbook(path, on_demand=True)
    except xlrd.biffh.XLRDError as e:
        msg = str(e)
        if "encrypt" in msg.lower() or "password" in msg.lower():
            raise ExtractError("パスワード保護されているため読めません")
        raise ExtractError(f"Excel(xls)を読めません: {msg}")
    units = []
    places = []
    try:
        for si in range(book.nsheets):
            sh = book.sheet_by_index(si)
            place = sh.name if sh.visibility == 0 else f"{sh.name}(非表示)"
            places.append(place)
            for r in range(sh.nrows):
                types = sh.row_types(r)
                vals = sh.row_values(r)
                for c, (ty, v) in enumerate(zip(types, vals)):
                    if ty in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                        continue
                    if ty == xlrd.XL_CELL_NUMBER:
                        v = str(int(v)) if float(v).is_integer() else repr(v)
                    elif ty == xlrd.XL_CELL_DATE:
                        try:
                            dt = xlrd.xldate.xldate_as_datetime(v, book.datemode)
                            v = dt.strftime("%Y/%m/%d %H:%M:%S").replace(" 00:00:00", "")
                        except Exception:
                            v = str(v)
                    elif ty == xlrd.XL_CELL_BOOLEAN:
                        v = "TRUE" if v else "FALSE"
                    elif ty == xlrd.XL_CELL_ERROR:
                        continue
                    v = str(v)
                    if v:
                        units.append(("F", place, r + 1, c + 1, f"{col_letter(c + 1)}{r + 1}", v))
            book.unload_sheet(si)
    finally:
        book.release_resources()
    # 図形・テキストボックス・コメントの文字 (xlrd は読まないので直接解析)
    from .oldoffice import xls_shape_units
    try:
        units.extend(xls_shape_units(path, places))
    except Exception:
        pass
    return units


# ---------------------------------------------------------------- Word (docx)
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
WPD = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
WPS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"


def _w_paragraph(p, shapes: list | None = None) -> tuple[str, tuple[int, int, int]]:
    """段落テキストと、段落内の (Wordが記録した改ページ数, 明示的な改ページ数, 段落先頭での記録改ページ数)。

    段落に置かれた図形・テキストボックスの文字は段落本文に混ぜず、shapes に (図形名, テキスト) で入れる。
    Word は同じ図形を新形式 (mc:Choice) と旧形式 VML (mc:Fallback) の2通りで保存するので Fallback は読まない。
    """
    parts = []
    cnt = [0, 0, 0]

    def rec(node, shape_name):
        for n in node:
            tag = n.tag
            if tag == W + "t":
                parts.append(n.text or "")
            elif tag == W + "tab":
                parts.append("\t")
            elif tag in (W + "br", W + "cr"):
                if n.get(W + "type") == "page":
                    cnt[1] += 1
                parts.append(" ")
            elif tag == W + "lastRenderedPageBreak":
                cnt[0] += 1
                if not "".join(parts).strip():
                    cnt[2] += 1  # 段落の先頭で改ページ → この段落は次のページ
            elif tag == W + "noBreakHyphen":
                parts.append("-")
            elif tag in (W + "delText", MC + "Fallback"):
                continue
            elif tag in (WPD + "anchor", WPD + "inline"):
                dp = n.find(WPD + "docPr")
                rec(n, (dp.get("name") if dp is not None else "") or shape_name)
            elif tag == WPS + "wsp":
                nv = n.find(WPS + "cNvPr")
                rec(n, (nv.get("name") if nv is not None else "") or shape_name)
            elif tag == W + "txbxContent":
                if shapes is not None:
                    for tp in n.iter(W + "p"):
                        t, _ = _w_paragraph(tp)
                        if t.strip():
                            shapes.append((shape_name, t))
            else:
                rec(n, shape_name)

    rec(p, "")
    return "".join(parts), (cnt[0], cnt[1], cnt[2])


def extract_docx(path: str) -> list[tuple]:
    z = _open_zip(path)
    units: list[tuple] = []
    with z:
        names = set(z.namelist())
        main = "word/document.xml"
        root_rels = _rels(z, "")  # _rels/.rels
        main = root_rels.get("type:officeDocument", main)
        if main not in names:
            raise ExtractError("Word 文書の本文が見つかりません")
        root = ET.fromstring(z.read(main))
        body = root.find(W + "body")
        state = {"para": 0, "r": 0, "e": 0, "table": 0}
        page_marks = []  # (units の index, 段落開始時点の rendered 累計, explicit 累計)

        def walk(container, table_ctx=None):
            for el in container:
                tag = el.tag
                if tag == W + "p":
                    state["para"] += 1
                    shapes = []
                    text, (rb, eb, lead) = _w_paragraph(el, shapes)
                    state["r"] += lead
                    rb -= lead
                    if text.strip():
                        place = "本文"
                        if table_ctx:
                            place = f"表{table_ctx[0]} ({table_ctx[1]}行,{table_ctx[2]}列)"
                        page_marks.append((len(units), state["r"], state["e"]))
                        units.append(("F", place, state["para"], None, "", text))
                    for sname, stext in shapes:
                        page_marks.append((len(units), state["r"], state["e"]))
                        units.append(("F", "図形", None, None,
                                      f"図形: {sname or 'テキストボックス'}（段落{state['para']}付近）", stext))
                    state["r"] += rb
                    state["e"] += eb
                elif tag == W + "tbl":
                    state["table"] += 1
                    tno = state["table"]
                    r = 0
                    for tr in el.findall(W + "tr"):
                        r += 1
                        c = 0
                        for tc in tr.findall(W + "tc"):
                            c += 1
                            walk(tc, (tno, r, c))
                elif tag in (W + "sdt", W + "sdtContent", W + "customXml", W + "smartTag",
                             W + "ins", W + "txbxContent"):
                    walk(el, table_ctx)
                else:
                    # テキストボックス等: 段落内の txbxContent は _w_paragraph に含まれる
                    if tag not in (W + "sectPr",):
                        walk(el, table_ctx)

        if body is not None:
            walk(body)
        # ページ番号はレイアウト依存のため、Word が保存時に記録した改ページ位置
        # (lastRenderedPageBreak) から概算する。無ければ明示的な改ページから数える。
        use_r = state["r"] > 0
        if state["r"] or state["e"]:
            for idx, rc, ec in page_marks:
                u = units[idx]
                page = f"約{(rc if use_r else ec) + 1}ページ"
                units[idx] = u[:4] + ((u[4] + " " + page) if u[4] else page,) + u[5:]
        # ヘッダー / フッター / 脚注 / 文末脚注 / コメント
        extra = [("header", "ヘッダー"), ("footer", "フッター"), ("footnotes", "脚注"),
                 ("endnotes", "文末脚注"), ("comments", "コメント")]
        for typ, label in extra:
            for i, part in enumerate(_rels_by_type(z, main, typ), 1):
                if part not in names:
                    continue
                r = ET.fromstring(z.read(part))
                skip = {id(x) for fb in r.iter(MC + "Fallback") for x in fb.iter(W + "p")}
                pn = 0
                for p in r.iter(W + "p"):
                    if id(p) in skip:
                        continue
                    pn += 1
                    t, _ = _w_paragraph(p)
                    if t.strip():
                        units.append(("F", f"{label}{i if typ in ('header', 'footer') else ''}",
                                      pn, None, "", t))
    return units


# ---------------------------------------------------------------- PowerPoint (pptx)
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _slide_lines(xml: bytes) -> list[str]:
    root = ET.fromstring(xml)
    lines = []
    for p in root.iter(A_NS + "p"):
        lines.append(_a_paragraph_text(p))
    # 末尾の空行を落とす
    while lines and not lines[-1].strip():
        lines.pop()
    return lines


def extract_pptx(path: str) -> list[tuple]:
    z = _open_zip(path)
    units: list[tuple] = []
    with z:
        names = set(z.namelist())
        root_rels = _rels(z, "")
        pres = root_rels.get("type:officeDocument", "ppt/presentation.xml")
        prels = _rels(z, pres)
        root = ET.fromstring(z.read(pres))
        slides = []
        lst = root.find(P_NS + "sldIdLst")
        if lst is not None:
            for s in lst:
                rid = next((v for k, v in s.attrib.items() if k.endswith("}id")), None)
                if rid in prels:
                    slides.append(prels[rid])
        for no, part in enumerate(slides, 1):
            if part not in names:
                continue
            lines = _slide_lines(z.read(part))
            if any(l.strip() for l in lines):
                units.append(("L", f"スライド{no}", 1, None, "", "\n".join(lines)))
            for npart in _rels_by_type(z, part, "notesSlide"):
                if npart in names:
                    nl = _slide_lines(z.read(npart))
                    # ノートにはスライド番号のプレースホルダが入るので数字だけの行は捨てない (行番号を保つ)
                    if any(l.strip() for l in nl):
                        units.append(("L", f"スライド{no} ノート", 1, None, "", "\n".join(nl)))
    return units


# ---------------------------------------------------------------- PDF
def extract_pdf(path: str) -> list[tuple]:
    import pypdfium2 as pdfium
    try:
        pdf = pdfium.PdfDocument(path)
    except pdfium.PdfiumError as e:
        msg = str(e)
        if "password" in msg.lower():
            raise ExtractError("パスワード保護されているため読めません")
        raise ExtractError(f"PDFを読めません: {msg}")
    units = []
    try:
        for i in range(len(pdf)):
            page = pdf[i]
            tp = page.get_textpage()
            text = tp.get_text_range()
            tp.close()
            page.close()
            text = text.replace("\r\n", "\n").replace("\r", "\n")
            if text.strip():
                units.append(("L", f"{i + 1}ページ", 1, None, "", text))
    finally:
        pdf.close()
    if not units:
        raise ExtractError("テキストが含まれていません (画像だけのPDFの可能性。OCRは未対応)")
    return units


# ---------------------------------------------------------------- 振り分け
def extract_units(path: str) -> list[tuple]:
    ext = os.path.splitext(path)[1].lower()
    if ext in EXCEL_NEW:
        return extract_xlsx(path)
    if ext in EXCEL_OLD:
        return extract_xls(path)
    if ext in WORD_NEW:
        return extract_docx(path)
    if ext in PPT_NEW:
        return extract_pptx(path)
    if ext in PDF_EXT:
        return extract_pdf(path)
    from . import oldoffice
    if ext in WORD_OLD:
        return oldoffice.extract_doc(path)
    if ext in PPT_OLD:
        return oldoffice.extract_ppt(path)
    raise ExtractError("未対応の形式です")
