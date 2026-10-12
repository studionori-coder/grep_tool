"""動作確認用のサンプルファイルを作る (開発用)。

python tests/make_samples.py <出力フォルダ>
旧形式 (.doc/.ppt/.xls) と PDF の作成には LibreOffice (soffice) と xlwt を使う。
"""
import os
import subprocess
import sys

import docx
import openpyxl
import pptx
from pptx.util import Inches

out = sys.argv[1] if len(sys.argv) > 1 else "samples"
os.makedirs(os.path.join(out, "sub", "deep"), exist_ok=True)
os.makedirs(os.path.join(out, "node_modules"), exist_ok=True)

# テキスト (いろいろな文字コード)
src = "# 基本設計書のサンプル\nline2: epic story\n\tインデント 設計 と EPIC\n最終行 epic\n"
open(os.path.join(out, "readme_utf8.txt"), "w", encoding="utf-8").write(src)
open(os.path.join(out, "sjis.csv"), "w", encoding="cp932").write("ID,名称,備考\n1,詳細設計,epic対応\n2,テスト,なし\n")
open(os.path.join(out, "sub", "euc.xml"), "w", encoding="euc_jp").write('<?xml version="1.0"?>\n<a>\n  <b>設計 epic</b>\n</a>\n')
open(os.path.join(out, "sub", "utf16.txt"), "w", encoding="utf-16").write("UTF16のファイル\nここに設計がある\n")
open(os.path.join(out, "sub", "deep", "Main.java"), "w", encoding="utf-8").write(
    "public class Main {\n    // epic: 設計参照\n    void epicMethod() {}\n}\n")
open(os.path.join(out, "node_modules", "x.js"), "w").write("epic in excluded dir\n")
open(os.path.join(out, "binary.dat2"), "wb").write(b"\x00\x01epic\x00\x00" * 100)

# Excel (xlsx)
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "機能一覧"
ws["A1"] = "No"
ws["B1"] = "機能名"
ws["B3"] = "ユーザー登録 (epic)"
ws["D10"] = "設計レビュー\n2行目にもepic"
ws2 = wb.create_sheet("隠しシート")
ws2["C5"] = "非表示シートの設計メモ"
ws2.sheet_state = "hidden"
wb.save(os.path.join(out, "機能一覧.xlsx"))

# Word (docx)
d = docx.Document()
d.sections[0].header.paragraphs[0].text = "ヘッダー: 基本設計書"
d.add_heading("1. 概要", 1)
d.add_paragraph("本書はシステムの基本設計を記述する。")
d.add_paragraph("epic の一覧は別紙。")
d.add_page_break()
t = d.add_table(rows=2, cols=2)
t.cell(0, 0).text = "項目"
t.cell(1, 1).text = "表の中の設計値"
d.add_paragraph("2ページ目の epic 本文")
d.save(os.path.join(out, "基本設計書.docx"))

# PowerPoint (pptx)
p = pptx.Presentation()
s1 = p.slides.add_slide(p.slide_layouts[1])
s1.shapes.title.text = "設計方針"
s1.placeholders[1].text = "ポイント1\nepic を分割する"
s1.notes_slide.notes_text_frame.text = "ノート: 設計の補足"
s2 = p.slides.add_slide(p.slide_layouts[5])
s2.shapes.title.text = "2枚目"
tb = s2.shapes.add_textbox(Inches(1), Inches(2), Inches(4), Inches(1))
tb.text_frame.text = "テキストボックスの epic"
p.save(os.path.join(out, "説明資料.pptx"))


def soffice(src, fmt, outdir):
    subprocess.run(["soffice", "--headless", "--convert-to", fmt, "--outdir", outdir, src],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


old = os.path.join(out, "旧形式")
os.makedirs(old, exist_ok=True)
soffice(os.path.join(out, "基本設計書.docx"), "doc", old)
soffice(os.path.join(out, "説明資料.pptx"), "ppt", old)
soffice(os.path.join(out, "機能一覧.xlsx"), "xls", old)
soffice(os.path.join(out, "基本設計書.docx"), "pdf", old)
print("created in", out)
