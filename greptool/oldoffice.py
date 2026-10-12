"""旧形式 Office (.doc / .ppt) のテキスト抽出。

Office をインストールしていなくても読めるよう、OLE 複合ファイルを直接解析する
([MS-DOC] のピーステーブル、[MS-PPT] のレコード構造)。
"""
from __future__ import annotations

import re
import struct

import olefile

from .extractors import ExtractError


def _open_ole(path: str) -> olefile.OleFileIO:
    if not olefile.isOleFile(path):
        raise ExtractError("ファイルが壊れているか、形式が拡張子と一致しません")
    return olefile.OleFileIO(path)


# ================================================================ Word 97-2003 (.doc)
def _doc_text(path: str) -> tuple[str, int]:
    ole = _open_ole(path)
    try:
        if not ole.exists("WordDocument"):
            raise ExtractError("Word 文書ではありません")
        wd = ole.openstream("WordDocument").read()
        ident, nfib = struct.unpack_from("<HH", wd, 0)
        if ident != 0xA5EC:
            raise ExtractError("Word 95 以前の形式には未対応です")
        flags = struct.unpack_from("<H", wd, 0x0A)[0]
        if flags & 0x0100:
            raise ExtractError("パスワード保護されているため読めません")
        table_name = "1Table" if flags & 0x0200 else "0Table"
        if not ole.exists(table_name):
            raise ExtractError("Word 文書の構造が壊れています")
        table = ole.openstream(table_name).read()

        pos = 32
        csw = struct.unpack_from("<H", wd, pos)[0]
        pos += 2 + csw * 2
        cslw = struct.unpack_from("<H", wd, pos)[0]
        rglw = pos + 2
        ccp = struct.unpack_from("<10i", wd, rglw + 3 * 4)  # ccpText, ccpFtn, ccpHdd, ccpMcr, ccpAtn, ccpEdn, ccpTxbx, ccpHdrTxbx ...
        ccp_text = ccp[0]
        ccp_all = ccp_text + sum(max(0, x) for x in (ccp[1], ccp[2], ccp[4], ccp[5], ccp[6], ccp[7]))
        ccp_all += max(0, ccp[3])
        pos = rglw + cslw * 4
        pos += 2  # cbRgFcLcb
        fc_clx, lcb_clx = struct.unpack_from("<II", wd, pos + 33 * 8)
        clx = table[fc_clx:fc_clx + lcb_clx]

        # Clx: Prc の並び (0x01) → Pcdt (0x02)
        i = 0
        while i < len(clx) and clx[i] == 0x01:
            cb = struct.unpack_from("<h", clx, i + 1)[0]
            i += 3 + cb
        if i >= len(clx) or clx[i] != 0x02:
            raise ExtractError("Word 文書のテキスト構造を解析できません")
        lcb = struct.unpack_from("<I", clx, i + 1)[0]
        plc = clx[i + 5:i + 5 + lcb]
        n = (lcb - 4) // 12
        cps = struct.unpack_from(f"<{n + 1}I", plc, 0)
        out = []
        total = 0
        for k in range(n):
            pcd_off = (n + 1) * 4 + k * 8
            fc_raw = struct.unpack_from("<I", plc, pcd_off + 2)[0]
            cnt = cps[k + 1] - cps[k]
            if cnt <= 0:
                continue
            if fc_raw & 0x40000000:
                fc = (fc_raw & 0x3FFFFFFF) // 2
                out.append(wd[fc:fc + cnt].decode("cp1252", "replace"))
            else:
                fc = fc_raw & 0x3FFFFFFF
                out.append(wd[fc:fc + cnt * 2].decode("utf-16-le", "replace"))
            total += cnt
            if total >= ccp_all:
                break
        text = "".join(out)
        # 各ストーリー (本文, 脚注, ヘッダー/フッター, マクロ, コメント, 文末脚注, テキストボックス, ヘッダー内テキストボックス)
        stories = [max(0, ccp[k]) for k in (0, 1, 2, 3, 4, 5, 6, 7)]
        return text[:ccp_all + max(0, ccp[3])], stories
    finally:
        ole.close()


def _strip_fields(text: str) -> str:
    """フィールド { コード \x14 結果 } から結果だけを残す"""
    if "\x13" not in text:
        return text
    out = []
    stack = []  # 各フィールドで「コード部分を読んでいる最中か」
    for ch in text:
        if ch == "\x13":
            stack.append(True)
        elif ch == "\x14":
            if stack:
                stack[-1] = False
        elif ch == "\x15":
            if stack:
                stack.pop()
        elif not any(stack):
            out.append(ch)
    return "".join(out)


_DOC_DROP = dict.fromkeys(map(ord, "\x01\x02\x03\x04\x05\x08"), None)
_DOC_DROP.update({0x0B: " ", 0x1E: "-", 0x1F: None, 0xA0: " "})


_STORY_NAMES = ["本文", "脚注", "ヘッダー/フッター", None, "コメント", "文末脚注", "図形", "図形(ヘッダー/フッター内)"]


def extract_doc(path: str) -> list[tuple]:
    try:
        raw, stories = _doc_text(path)
    except ExtractError:
        raise
    except Exception as e:  # 解析失敗
        raise ExtractError(f"Word(doc)を解析できません: {e}")
    units = []
    pos = 0
    for idx, length in enumerate(stories):
        part_text = raw[pos:pos + length]
        pos += length
        place = _STORY_NAMES[idx]
        if place is None or not part_text:
            continue
        text = _strip_fields(part_text)
        main = idx == 0
        has_pb = main and "\x0c" in text
        page = 1
        seen = set()
        # 段落区切り: \r (段落), \x07 (表のセル/行の終わり)
        for para, line in enumerate(re.split("[\r\x07]", text), 1):
            pb = line.count("\x0c")
            line = line.replace("\x0c", " ").translate(_DOC_DROP)
            if line.strip():
                if main:
                    label = f"{page}ページ以降" if has_pb else ""
                    units.append(("F", place, para, None, label, line))
                elif line not in seen:  # ヘッダー等は奇数/偶数/先頭ページ用に重複して入る
                    seen.add(line)
                    if place.startswith("図形"):
                        units.append(("F", place, None, None, "図形", line))
                    else:
                        units.append(("F", place, para, None, "", line))
            page += pb
    return units


# ================================================================ PowerPoint 97-2003 (.ppt)
RT_DOCUMENT = 0x03E8
RT_SLIDE = 0x03EE
RT_NOTES = 0x03F0
RT_NOTES_ATOM = 0x03F1
RT_SLIDE_PERSIST = 0x03F3
RT_MAIN_MASTER = 0x03F8
RT_SLIDE_LIST = 0x0FF0
RT_TEXT_CHARS = 0x0FA0
RT_TEXT_BYTES = 0x0FA8
RT_CSTRING = 0x0FBA
RT_PERSIST_DIR = 0x1772
RT_CRYPT = 0x2F14


def _iter_records(data: bytes, start: int, end: int, depth=0):
    """(offset, recVer, recInstance, recType, body_start, body_end) を深さ優先で列挙"""
    pos = start
    while pos + 8 <= end:
        vi, rtype, rlen = struct.unpack_from("<HHI", data, pos)
        ver, inst = vi & 0x0F, vi >> 4
        bstart, bend = pos + 8, min(pos + 8 + rlen, end)
        yield pos, ver, inst, rtype, bstart, bend, depth
        if ver == 0x0F and depth < 30:
            yield from _iter_records(data, bstart, bend, depth + 1)
        pos = pos + 8 + rlen


def _ppt_text_atom(data, rtype, b, e) -> str:
    if rtype == RT_TEXT_CHARS:
        return data[b:e].decode("utf-16-le", "replace")
    return data[b:e].decode("latin-1")


def extract_ppt(path: str) -> list[tuple]:
    ole = _open_ole(path)
    try:
        if not ole.exists("PowerPoint Document"):
            if ole.exists("EncryptedSummary") or ole.exists("EncryptionInfo"):
                raise ExtractError("パスワード保護されているため読めません")
            raise ExtractError("PowerPoint 文書ではありません")
        data = ole.openstream("PowerPoint Document").read()
    finally:
        ole.close()

    try:
        return _parse_ppt(data)
    except ExtractError:
        raise
    except Exception as e:
        raise ExtractError(f"PowerPoint(ppt)を解析できません: {e}")


def _parse_ppt(data: bytes) -> list[tuple]:
    persist: dict[int, int] = {}
    # top-level をまず走査し、persist ディレクトリと最後の DocumentContainer を得る
    doc_rec = None
    top = []
    pos = 0
    while pos + 8 <= len(data):
        vi, rtype, rlen = struct.unpack_from("<HHI", data, pos)
        top.append((pos, rtype, pos + 8, min(pos + 8 + rlen, len(data))))
        if rtype == RT_DOCUMENT:
            doc_rec = (pos + 8, min(pos + 8 + rlen, len(data)))
        elif rtype == RT_PERSIST_DIR:
            p, e = pos + 8, pos + 8 + rlen
            while p + 4 <= e:
                hdr = struct.unpack_from("<I", data, p)[0]
                pid, cnt = hdr & 0xFFFFF, hdr >> 20
                p += 4
                for k in range(cnt):
                    if p + 4 > e:
                        break
                    persist[pid + k] = struct.unpack_from("<I", data, p)[0]
                    p += 4
        elif rtype == RT_CRYPT:
            raise ExtractError("パスワード保護されているため読めません")
        if rlen == 0 and rtype == 0:
            break
        pos += 8 + rlen
    if doc_rec is None:
        raise ExtractError("PowerPoint 文書の構造を解析できません")

    # スライド一覧 (表示順) と、プレースホルダのテキスト
    slide_texts: dict[int, list[str]] = {}
    notes_texts: dict[int, list[str]] = {}
    offset_to_slide: dict[int, int] = {}
    slideid_to_no: dict[int, int] = {}
    n_slides = 0
    current_list_inst = None
    cur = None
    list_end = -1
    for off, ver, inst, rtype, b, e, depth in _iter_records(data, *doc_rec):
        if off >= list_end:
            current_list_inst = None
        if rtype == RT_SLIDE_LIST:
            current_list_inst = inst
            list_end = e
            cur = None
            continue
        if current_list_inst is None:
            continue
        if rtype == RT_SLIDE_PERSIST and current_list_inst == 0:
            n_slides += 1
            cur = n_slides
            pid, _flags, _ct, slide_id = struct.unpack_from("<IIiI", data, b)
            if pid in persist:
                offset_to_slide[persist[pid]] = cur
            slideid_to_no[slide_id] = cur
        elif rtype in (RT_TEXT_CHARS, RT_TEXT_BYTES) and current_list_inst == 0 and cur:
            slide_texts.setdefault(cur, []).append(_ppt_text_atom(data, rtype, b, e))

    # スライド/ノートのコンテナ内のテキスト (プレースホルダ以外の図形・テキストボックス)
    for off, rtype, b, e in top:
        if rtype == RT_SLIDE and off in offset_to_slide:
            no = offset_to_slide[off]
            lst = slide_texts.setdefault(no, [])
            for _o, _v, _i, rt, bb, ee, _d in _iter_records(data, b, e):
                if rt in (RT_TEXT_CHARS, RT_TEXT_BYTES):
                    t = _ppt_text_atom(data, rt, bb, ee)
                    if t not in lst:
                        lst.append(t)
        elif rtype == RT_NOTES:
            no = None
            texts = []
            for _o, _v, _i, rt, bb, ee, _d in _iter_records(data, b, e):
                if rt == RT_NOTES_ATOM and ee - bb >= 4:
                    sid = struct.unpack_from("<I", data, bb)[0]
                    no = slideid_to_no.get(sid)
                elif rt in (RT_TEXT_CHARS, RT_TEXT_BYTES):
                    texts.append(_ppt_text_atom(data, rt, bb, ee))
            if no and texts:
                lst = notes_texts.setdefault(no, [])
                for t in texts:
                    if t.strip() != "*" and t not in lst:  # "*" はスライド番号の差し込み
                        lst.append(t)

    units = []
    for no in sorted(set(slide_texts) | set(notes_texts)):
        for src, place in ((slide_texts, f"スライド{no}"), (notes_texts, f"スライド{no} ノート")):
            texts = src.get(no)
            if not texts:
                continue
            body = "\n".join(t.replace("\r", "\n").replace("\x0b", " ") for t in texts)
            if body.strip():
                units.append(("L", place, 1, None, "", body))
    return units


# ================================================================ Excel 97-2003 (.xls) の図形・コメント
def xls_shape_units(path: str, places: list[str]) -> list[tuple]:
    """BIFF8 の TXO レコード (テキストボックス・図形・セルのコメント) から文字を取り出す。
    places: シート順の表示名 (xlrd で得たもの)。セル値は xlrd が読むので、ここでは図形とコメントだけ。"""
    try:
        ole = _open_ole(path)
    except ExtractError:
        return []
    try:
        name = "Workbook" if ole.exists("Workbook") else ("Book" if ole.exists("Book") else None)
        if not name:
            return []
        data = ole.openstream(name).read()
    finally:
        ole.close()

    recs = []
    pos = 0
    n = len(data)
    while pos + 4 <= n:
        rt, ln = struct.unpack_from("<HH", data, pos)
        recs.append((pos, rt, pos + 4, min(pos + 4 + ln, n)))
        pos += 4 + ln

    bof_offsets = []  # BOUNDSHEET の lbPlyPos (シートの BOF 位置) をシート順に
    for off, rt, b, e in recs:
        if rt == 0x0085 and e - b >= 4:
            bof_offsets.append(struct.unpack_from("<I", data, b)[0])
    sheet_of_bof = {o: i for i, o in enumerate(bof_offsets)}

    units = []
    sheet = None
    anchor = ""
    obj_type, obj_id = None, None
    notes_text: dict[int, str] = {}
    i = 0
    while i < len(recs):
        off, rt, b, e = recs[i]
        if rt == 0x0809:  # BOF
            if off in sheet_of_bof:
                sheet = sheet_of_bof[off]
                notes_text = {}
        elif rt == 0x00EC and sheet is not None:  # MSODRAWING: 図形の置き場所 (ClientAnchor) を拾う
            k = data.rfind(b"\x10\xf0\x12\x00\x00\x00", b, e)
            if k >= 0 and k + 6 + 18 <= e:
                col, _dx, row = struct.unpack_from("<HHH", data, k + 8)
                from .extractors import col_letter
                anchor = f"（{col_letter(col + 1)}{row + 1}付近）"
        elif rt == 0x005D and e - b >= 8:  # OBJ (ftCmo)
            obj_type, obj_id = struct.unpack_from("<HH", data, b + 4)
        elif rt == 0x01B6 and sheet is not None and e - b >= 12:  # TXO
            cch = struct.unpack_from("<H", data, b + 10)[0]
            chars = []
            got = 0
            j = i + 1
            while got < cch and j < len(recs) and recs[j][1] == 0x003C:
                _, _, cb, ce = recs[j]
                high = data[cb] & 0x01
                body = data[cb + 1:ce]
                if high:
                    take = min(cch - got, len(body) // 2)
                    chars.append(body[:take * 2].decode("utf-16-le", "replace"))
                else:
                    take = min(cch - got, len(body))
                    chars.append(body[:take].decode("latin-1"))
                got += take
                j += 1
            text = "".join(chars).replace("\r\n", "\n").replace("\r", "\n")
            if text.strip() and sheet < len(places):
                if obj_type == 0x19:  # セルのコメント → 後の NOTE レコードでセル位置が分かる
                    notes_text[obj_id] = text
                else:
                    units.append(("F", places[sheet], None, None, f"図形{anchor}", text))
            i = j - 1
        elif rt == 0x001C and sheet is not None and e - b >= 8:  # NOTE
            row, col, _flags, oid = struct.unpack_from("<HHHH", data, b)
            if oid in notes_text and sheet < len(places):
                from .extractors import col_letter
                units.append(("F", places[sheet], row + 1, col + 1, f"{col_letter(col + 1)}{row + 1} (コメント)",
                              notes_text.pop(oid)))
        i += 1
    return units
