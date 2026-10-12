"""画面 (PySide6)。"""
from __future__ import annotations

import datetime as dt
import html
import os
import re
import sys
import threading

from PySide6.QtCore import (QAbstractTableModel, QModelIndex, QSettings, QSize, QSortFilterProxyModel,
                            Qt, QThread, QTimer, Signal, QRect)
from PySide6.QtGui import QAction, QColor, QFont, QGuiApplication, QKeySequence, QPalette, QTextDocument, QTextOption
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QButtonGroup, QCheckBox, QComboBox, QDialog,
                               QDialogButtonBox, QFileDialog, QFormLayout, QGridLayout, QGroupBox, QHBoxLayout,
                               QHeaderView, QLabel, QLineEdit, QMainWindow, QMenu, QMessageBox, QProgressBar,
                               QProgressDialog, QPushButton, QRadioButton, QSplitter, QStyle, QStyledItemDelegate,
                               QTableView, QTabWidget, QTextBrowser, QVBoxLayout, QWidget)

from . import __version__, opener
from .cache import ExtractCache, default_cache_path
from .engine import SearchEngine
from .extractors import ENCODING_CHOICES, ENCODING_DISPLAY, KIND_LABEL, decode_text

APP = "GrepTool"
HIST_MAX = 20

LEGEND = ("<b>行・列の見方</b>　"
          "テキスト：行番号／行頭からの文字位置　｜　Excel：セルの行／列（位置にセル番地）　｜　"
          "Word：段落番号／段落内の文字位置　｜　PowerPoint・PDF：スライド・ページ内の行／文字位置　｜　"
          "<span style='color:#888'>－</span>：図形など行・列が無い箇所")


def fmt_secs(s: float | None) -> str:
    if s is None:
        return "計算中…"
    s = int(round(s))
    if s < 60:
        return f"約{max(1, s)}秒"
    m, s = divmod(s, 60)
    if m < 60:
        return f"約{m}分{s:02d}秒"
    h, m = divmod(m, 60)
    return f"約{h}時間{m:02d}分"


def fmt_elapsed(s: float) -> str:
    s = int(s)
    return f"{s // 60}:{s % 60:02d}"


def fmt_size(n: int) -> str:
    for u in ("B", "KB", "MB", "GB"):
        if n < 1024 or u == "GB":
            return f"{n:,.0f}{u}" if u == "B" else f"{n:,.1f}{u}"
        n /= 1024


# ====================================================================== モデル
class HitModel(QAbstractTableModel):
    COLS = ["No", "ファイル", "種別", "場所", "行", "列", "位置", "該当箇所", "件数", "フォルダ", "更新日時"]
    TIPS = {
        1: "クリックすると該当箇所を開きます",
        3: "Excel=シート名 / Word=本文・表(行,列)・ヘッダー等 / PowerPoint=スライド番号 / PDF=ページ番号",
        4: "テキスト=行番号 / Excel=セルの行 / Word=段落番号 / PowerPoint・PDF=スライド・ページ内の行",
        5: "テキスト・Word・PowerPoint・PDF=文字位置(何文字目) / Excel=セルの列番号",
        6: "Excel=セル番地 / Word=ページ(目安) / 図形の名前 など",
        8: "「該当行」モードで、同じ行(セル)に複数ヒットした場合の件数",
    }
    C_NO, C_FILE, C_KIND, C_PLACE, C_LINE, C_COL, C_LABEL, C_CTX, C_CNT, C_DIR, C_TIME = range(11)

    def __init__(self):
        super().__init__()
        self.hits: list[tuple] = []
        self.files = []

    def reset(self, files=None):
        self.beginResetModel()
        self.hits = []
        self.files = files or []
        self.endResetModel()

    def append(self, hits):
        if not hits:
            return
        n = len(self.hits)
        self.beginInsertRows(QModelIndex(), n, n + len(hits) - 1)
        self.hits.extend(hits)
        self.endInsertRows()

    def finalize_order(self):
        """検索完了後、ファイルの列挙順(フォルダ・名前順)に並べ直す"""
        self.layoutAboutToBeChanged.emit()
        self.hits.sort(key=lambda h: h[0])  # 安定ソートなのでファイル内の順序は保たれる
        self.layoutChanged.emit()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.hits)

    def columnCount(self, parent=QModelIndex()):
        return len(self.COLS)

    def headerData(self, section, orient, role=Qt.DisplayRole):
        if orient == Qt.Horizontal:
            if role == Qt.DisplayRole:
                return self.COLS[section]
            if role == Qt.ToolTipRole:
                return self.TIPS.get(section)
        elif role == Qt.DisplayRole:
            return None
        return None

    def file_of(self, h):
        return self.files[h[0]]

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        h = self.hits[index.row()]
        c = index.column()
        if role == Qt.DisplayRole:
            f = self.files[h[0]]
            if c == self.C_NO:
                return index.row() + 1
            if c == self.C_FILE:
                return os.path.basename(f.path)
            if c == self.C_KIND:
                return KIND_LABEL.get(h[1], h[1])
            if c == self.C_PLACE:
                return h[2]
            if c == self.C_LINE:
                return h[3] if h[3] is not None else "－"
            if c == self.C_COL:
                return h[4] if h[4] is not None else "－"
            if c == self.C_LABEL:
                return h[5]
            if c == self.C_CTX:
                return h[7]
            if c == self.C_CNT:
                return h[10]
            if c == self.C_DIR:
                return os.path.dirname(f.path)
            if c == self.C_TIME:
                return dt.datetime.fromtimestamp(f.mtime).strftime("%Y/%m/%d %H:%M")
        elif role == Qt.ToolTipRole:
            f = self.files[h[0]]
            if c == self.C_FILE:
                return f"{f.path}\n（クリックで該当箇所を開く）"
            if c in (self.C_LINE, self.C_COL) and h[3] is None:
                return "図形・テキストボックス内のため行・列はありません"
            if c == self.C_CTX:
                return h[7]
            if c == self.C_DIR:
                return os.path.dirname(f.path)
        elif role == Qt.TextAlignmentRole:
            if c in (self.C_NO, self.C_LINE, self.C_COL, self.C_CNT):
                return int(Qt.AlignRight | Qt.AlignVCenter)
        elif role == Qt.UserRole:  # ソート用
            f = self.files[h[0]]
            if c == self.C_NO:
                return index.row()
            if c == self.C_FILE:
                return (os.path.basename(f.path).lower(), h[0])
            if c in (self.C_LINE, self.C_COL):
                v = h[3] if c == self.C_LINE else h[4]
                return (h[0], h[2], v if v is not None else -1)
            if c == self.C_CNT:
                return h[10]
            if c == self.C_TIME:
                return f.mtime
            return str(self.data(index, Qt.DisplayRole))
        return None


class SortProxy(QSortFilterProxyModel):
    def __init__(self):
        super().__init__()
        self.setSortRole(Qt.UserRole)
        self.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self._text = ""

    def set_text(self, t):
        self._text = t.lower()
        self.invalidateFilter()

    def filterAcceptsRow(self, row, parent):
        if not self._text:
            return True
        m = self.sourceModel()
        h = m.hits[row]
        f = m.files[h[0]]
        return self._text in f.path.lower() or self._text in h[7].lower() or self._text in (h[2] or "").lower()

    def lessThan(self, a, b):
        va, vb = a.data(Qt.UserRole), b.data(Qt.UserRole)
        try:
            return va < vb
        except TypeError:
            return str(va) < str(vb)


class FileModel(QAbstractTableModel):
    """ファイル別 / 読めなかったファイル 用"""

    def __init__(self, errors_mode=False):
        super().__init__()
        self.errors_mode = errors_mode
        self.rows = []  # FileInfo または (path, message) のタプル
        if errors_mode:
            self.COLS = ["ファイル / フォルダ", "種別", "状態", "理由", "フォルダ"]
        else:
            self.COLS = ["ファイル", "種別", "ヒット件数", "文字コード", "サイズ", "更新日時", "フォルダ", "備考"]

    def reset(self):
        self.beginResetModel()
        self.rows = []
        self.endResetModel()

    def append(self, items):
        if not items:
            return
        n = len(self.rows)
        self.beginInsertRows(QModelIndex(), n, n + len(items) - 1)
        self.rows.extend(items)
        self.endInsertRows()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return len(self.COLS)

    def headerData(self, s, o, role=Qt.DisplayRole):
        if o == Qt.Horizontal and role == Qt.DisplayRole:
            return self.COLS[s]
        return None

    def path_at(self, row):
        r = self.rows[row]
        return r[0] if isinstance(r, tuple) else r.path

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        r = self.rows[index.row()]
        c = index.column()
        if role in (Qt.DisplayRole, Qt.UserRole):
            if isinstance(r, tuple):  # フォルダ列挙のエラー
                return [os.path.basename(r[0]) or r[0], "フォルダ", "エラー", r[1], os.path.dirname(r[0])][c]
            if self.errors_mode:
                return [os.path.basename(r.path), KIND_LABEL.get(r.kind, r.kind),
                        "エラー" if r.status == "error" else "スキップ", r.message, os.path.dirname(r.path)][c]
            if c == 0:
                return os.path.basename(r.path)
            if c == 1:
                return KIND_LABEL.get(r.kind, r.kind)
            if c == 2:
                return r.hits
            if c == 3:
                return ENCODING_DISPLAY.get(r.encoding, r.encoding)
            if c == 4:
                return r.size if role == Qt.UserRole else fmt_size(r.size)
            if c == 5:
                return r.mtime if role == Qt.UserRole else dt.datetime.fromtimestamp(r.mtime).strftime("%Y/%m/%d %H:%M")
            if c == 6:
                return os.path.dirname(r.path)
            if c == 7:
                return "ヒットが多いため途中で打ち切り" if r.truncated else ("キャッシュから検索" if r.status == "cached" else "")
        if role == Qt.ToolTipRole:
            return self.path_at(index.row())
        if role == Qt.TextAlignmentRole and not self.errors_mode and c in (2, 4):
            return int(Qt.AlignRight | Qt.AlignVCenter)
        return None


# ====================================================================== 描画
class LinkDelegate(QStyledItemDelegate):
    """ファイル名をリンク風に描く"""

    def paint(self, painter, option, index):
        self.initStyleOption(option, index)
        option.font.setUnderline(True)
        if not (option.state & QStyle.State_Selected):
            option.palette.setColor(QPalette.Text, QColor("#0563C1"))
        super().paint(painter, option, index)


class HighlightDelegate(QStyledItemDelegate):
    """該当箇所のヒット部分を黄色で強調して描く"""

    def __init__(self, model_getter, parent=None):
        super().__init__(parent)
        self.model_getter = model_getter
        self.doc = QTextDocument()
        opt = QTextOption()
        opt.setWrapMode(QTextOption.NoWrap)
        self.doc.setDefaultTextOption(opt)
        self.doc.setDocumentMargin(2)

    def paint(self, painter, option, index):
        self.initStyleOption(option, index)
        src = index.model().mapToSource(index) if hasattr(index.model(), "mapToSource") else index
        h = self.model_getter().hits[src.row()]
        ctx, hs, he = h[7], h[8], h[9]
        style = option.widget.style() if option.widget else QApplication.style()
        text_bak = option.text
        option.text = ""
        style.drawControl(QStyle.CE_ItemViewItem, option, painter, option.widget)
        option.text = text_bak
        selected = bool(option.state & QStyle.State_Selected)
        color = option.palette.color(QPalette.HighlightedText if selected else QPalette.Text).name()
        body = (html.escape(ctx[:hs]) + "<span style='background:#ffe066;color:#000;font-weight:bold'>" +
                html.escape(ctx[hs:he]) + "</span>" + html.escape(ctx[he:]))
        self.doc.setDefaultFont(option.font)
        self.doc.setHtml(f"<div style='color:{color};white-space:pre'>{body}</div>")
        painter.save()
        r = option.rect
        painter.setClipRect(r)
        # ヒット位置が見えるように、必要なら左へずらす
        shift = 0
        if hs > 40:
            pre_doc = QTextDocument()
            pre_doc.setDefaultFont(option.font)
            pre_doc.setPlainText(ctx[:hs])
            pre_w = pre_doc.idealWidth()
            if pre_w > r.width() * 0.6:
                shift = pre_w - r.width() * 0.3
        painter.translate(r.left() - shift, r.top() + (r.height() - self.doc.size().height()) / 2)
        self.doc.drawContents(painter)
        painter.restore()


# ====================================================================== 検索スレッド
class SearchThread(QThread):
    sig_enum = Signal(int)
    sig_started = Signal(object, int)
    sig_progress = Signal(int, int, float, object, str, int)
    sig_results = Signal(object, object)
    sig_finished = Signal(object, object)
    sig_error = Signal(str)

    def __init__(self, engine: SearchEngine, opts: dict):
        super().__init__()
        self.engine = engine
        self.opts = opts
        self.cancel = threading.Event()

    # engine から呼ばれるコールバック
    def enumerating(self, n):
        self.sig_enum.emit(n)

    def started(self, files, total_bytes):
        self.sig_started.emit(files, total_bytes)

    def progress(self, done, total, ratio, eta, cur, hits):
        self.sig_progress.emit(done, total, ratio, eta, cur, hits)

    def results(self, files, hits):
        self.sig_results.emit(list(files), list(hits))

    def run(self):
        try:
            files, summ = self.engine.run(self.opts, self, self.cancel)
            self.sig_finished.emit(files, summ)
        except re.error as e:
            self.sig_error.emit(f"正規表現に誤りがあります: {e}")
        except Exception as e:
            self.sig_error.emit(f"検索中にエラーが発生しました: {type(e).__name__}: {e}")


class ExportThread(QThread):
    sig_progress = Signal(int, int)
    sig_done = Signal(str)

    def __init__(self, path, hits, files, cond, summary):
        super().__init__()
        self.args = (path, hits, files, cond, summary)

    def run(self):
        from .excel_export import export
        try:
            export(*self.args, progress=lambda i, n: self.sig_progress.emit(i, n))
            self.sig_done.emit("")
        except PermissionError:
            self.sig_done.emit("保存できませんでした。同じ名前のファイルを Excel で開いていないか確認してください。")
        except Exception as e:
            self.sig_done.emit(f"保存できませんでした: {e}")


# ====================================================================== 設定ダイアログ
class SettingsDialog(QDialog):
    def __init__(self, parent, editor: str, cache: ExtractCache | None):
        super().__init__(parent)
        self.setWindowTitle("設定")
        self.setMinimumWidth(640)
        lay = QVBoxLayout(self)
        g = QGroupBox("テキストファイルを開くエディタ（行・桁を指定して開きます）")
        fl = QFormLayout(g)
        self.preset = QComboBox()
        for name, cmd in opener.EDITOR_PRESETS:
            self.preset.addItem(name, cmd)
        self.cmd = QLineEdit(editor)
        self.cmd.setPlaceholderText('例: "C:\\Program Files (x86)\\sakura\\sakura.exe" -Y={line} -X={col} "{path}"')
        self.preset.currentIndexChanged.connect(lambda i: self.cmd.setText(self.preset.itemData(i)))
        fl.addRow("プリセット", self.preset)
        fl.addRow("コマンド", self.cmd)
        fl.addRow("", QLabel("{path}=ファイル, {line}=行, {col}=桁 に置き換わります。空欄なら既定のアプリで開きます。"))
        lay.addWidget(g)

        g2 = QGroupBox("抽出キャッシュ（Office・PDF の2回目以降の検索を高速化）")
        v2 = QVBoxLayout(g2)
        self.cache = cache
        self.cache_label = QLabel()
        self._update_cache_label()
        btn = QPushButton("キャッシュを削除")
        btn.clicked.connect(self._clear_cache)
        h2 = QHBoxLayout()
        h2.addWidget(self.cache_label, 1)
        h2.addWidget(btn)
        v2.addLayout(h2)
        v2.addWidget(QLabel(f"保存場所: {default_cache_path()}"))
        lay.addWidget(g2)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def _update_cache_label(self):
        if self.cache:
            n, size = self.cache.stats()
            self.cache_label.setText(f"{n:,} ファイル分 / {fmt_size(size)}")
        else:
            self.cache_label.setText("（キャッシュなし）")

    def _clear_cache(self):
        if self.cache and QMessageBox.question(self, "確認", "キャッシュを削除しますか？\n（次回の検索は少し遅くなります）") \
                == QMessageBox.Yes:
            self.cache.clear()
            self._update_cache_label()


HELP_HTML = """
<h3>使い方</h3>
<ol>
<li><b>条件</b>に探したい文字列を入れ、<b>フォルダ</b>を選んで <b>検索(F)</b>（または Enter）。</li>
<li>結果の<b>ファイル名（青字）をクリック</b>すると、該当箇所を開きます（Excel はシートとセル、Word は段落、
PowerPoint はスライドへ移動。Windows で pywin32 が入っている場合）。</li>
<li><b>Excelに出力(E)</b> で結果を保存できます。Excel 上のファイル名もリンクになっています。</li>
</ol>
<h3>対象ファイル</h3>
<ul>
<li>Excel: .xlsx .xlsm .xls（セル・コメント・図形/テキストボックス内の文字）</li>
<li>Word: .docx .docm .doc（本文・表・図形/テキストボックス・ヘッダー/フッター・脚注・コメント）</li>
<li>PowerPoint: .pptx .pptm .ppt（スライド・ノート）</li>
<li>PDF: .pdf（文字情報があるもの。画像だけのPDFは対象外）</li>
<li>テキスト: 上記以外の拡張子すべて（.txt .csv .xml .java .sql .html など）。中身がバイナリのものは自動でスキップ。</li>
</ul>
<h3>行・列の見方</h3>
<table border=1 cellspacing=0 cellpadding=4>
<tr><th>種別</th><th>場所</th><th>行</th><th>列</th><th>位置</th></tr>
<tr><td>テキスト</td><td>－</td><td>行番号</td><td>行頭からの文字位置</td><td>－</td></tr>
<tr><td>Excel</td><td>シート名</td><td>セルの行</td><td>セルの列(数字)</td><td>セル番地 (B12)</td></tr>
<tr><td>Word</td><td>本文 / 表n(行,列) / ヘッダー等</td><td>段落番号</td><td>段落内の文字位置</td><td>ページ(目安)</td></tr>
<tr><td>PowerPoint</td><td>スライドn / ノート</td><td>スライド内の行</td><td>文字位置</td><td>－</td></tr>
<tr><td>PDF</td><td>nページ</td><td>ページ内の行</td><td>文字位置</td><td>－</td></tr>
</table>
<p>図形・テキストボックスの中の文字は行・列が無いため「－」と表示し、位置欄に図形名を出します。<br>
Word のページ番号は、Word が保存時に記録した改ページ位置から求めた目安です。</p>
<h3>速さについて</h3>
<p>複数の CPU コアで並列に検索します。Office・PDF から取り出した文字は<b>抽出キャッシュ</b>に保存されるので、
2回目以降の検索は大幅に速くなります（ファイルが更新されたら自動で読み直します）。</p>
<h3>正規表現</h3>
<p>Python の正規表現です。例: <code>設計|仕様</code>、<code>ID-\\d{4}</code></p>
"""


# ====================================================================== メイン画面
class MainWindow(QMainWindow):
    def __init__(self, engine: SearchEngine, initial_folder: str | None = None):
        super().__init__()
        self.engine = engine
        self.settings = QSettings(APP, APP)
        self.thread: SearchThread | None = None
        self.files = []
        self.summary = None
        self.last_opts = None
        try:
            self.cache = ExtractCache(default_cache_path())
        except Exception:
            self.cache = None
        self.editor = self.settings.value("editor", None)
        if self.editor is None:
            self.editor = opener.detect_editor()

        self.setWindowTitle(f"Grep ツール  {__version__}")
        self.resize(1280, 860)
        self._build_ui()
        self._load_settings()
        if initial_folder:
            self.folder.setCurrentText(initial_folder)
        self.setAcceptDrops(True)
        QTimer.singleShot(0, self.cond.setFocus)

    # ------------------------------------------------------------------ UI 構築
    def _combo(self, editable=True):
        c = QComboBox()
        c.setEditable(editable)
        c.setInsertPolicy(QComboBox.NoInsert)
        c.setMaxCount(HIST_MAX)
        c.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        return c

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(10, 8, 10, 6)

        # ---- 検索条件
        box = QGroupBox("検索条件")
        grid = QGridLayout(box)
        grid.setHorizontalSpacing(8)
        grid.setVerticalSpacing(5)

        self.cond = self._combo()
        self.cond.lineEdit().setPlaceholderText("検索する文字列")
        self.cond.lineEdit().returnPressed.connect(self.start_search)
        self.include = self._combo()
        self.include.lineEdit().setPlaceholderText("*.*（すべて）　例: *.xlsx;*.docx;*.java")
        self.folder = self._combo()
        self.folder.lineEdit().setPlaceholderText("検索するフォルダ（; 区切りで複数可。ここへドラッグ＆ドロップも可）")
        self.ex_files = self._combo()
        self.ex_files.lineEdit().setPlaceholderText("例: *.log;*.min.js")
        self.ex_dirs = self._combo()
        self.ex_dirs.lineEdit().setPlaceholderText("例: .git;.svn;node_modules;bin;obj")

        def lbl(text, buddy):
            l = QLabel(text)
            l.setBuddy(buddy)
            return l

        btn_browse = QPushButton("…")
        btn_browse.setFixedWidth(32)
        btn_browse.setToolTip("フォルダを選ぶ")
        btn_browse.clicked.connect(self.browse_folder)
        btn_up = QPushButton("上フォルダ(&B)")
        btn_up.clicked.connect(self.parent_folder)

        grid.addWidget(lbl("条件(&N)", self.cond), 0, 0)
        grid.addWidget(self.cond, 0, 1, 1, 2)
        grid.addWidget(lbl("ファイル(&T)", self.include), 1, 0)
        grid.addWidget(self.include, 1, 1, 1, 2)
        grid.addWidget(lbl("フォルダ(&O)", self.folder), 2, 0)
        grid.addWidget(self.folder, 2, 1)
        grid.addWidget(btn_browse, 2, 2)
        grid.addWidget(lbl("除外ファイル", self.ex_files), 3, 0)
        grid.addWidget(self.ex_files, 3, 1, 1, 2)
        grid.addWidget(lbl("除外フォルダ", self.ex_dirs), 4, 0)
        grid.addWidget(self.ex_dirs, 4, 1, 1, 2)

        # 右側のボタン列
        self.btn_search = QPushButton("検索(&F)")
        self.btn_search.setDefault(True)
        self.btn_search.setMinimumHeight(40)
        f = self.btn_search.font()
        f.setBold(True)
        f.setPointSize(f.pointSize() + 1)
        self.btn_search.setFont(f)
        self.btn_search.setStyleSheet(
            "QPushButton{background:#1f6feb;color:white;border-radius:4px;padding:4px 18px}"
            "QPushButton:hover{background:#2f7ff8}QPushButton:disabled{background:#9db8e0}")
        self.btn_search.clicked.connect(self.start_search)
        self.btn_cancel = QPushButton("キャンセル(&X)")
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self.cancel_search)
        btn_help = QPushButton("ヘルプ(&H)")
        btn_help.clicked.connect(self.show_help)
        side = QVBoxLayout()
        side.addWidget(self.btn_search)
        side.addWidget(self.btn_cancel)
        side.addWidget(btn_up)
        side.addWidget(btn_help)
        side.addStretch(1)
        grid.addLayout(side, 0, 3, 5, 1)
        grid.setColumnStretch(1, 1)

        # オプション
        opt_row = QHBoxLayout()
        g1 = QGroupBox("検索方法")
        v1 = QVBoxLayout(g1)
        self.chk_word = QCheckBox("単語単位で探す(&W)")
        self.chk_case = QCheckBox("英大文字と小文字を区別する(&C)")
        self.chk_regex = QCheckBox("正規表現(&E)")
        for w in (self.chk_word, self.chk_case, self.chk_regex):
            v1.addWidget(w)
        v1.addStretch(1)

        g2 = QGroupBox("対象")
        v2 = QGridLayout(g2)
        self.chk_sub = QCheckBox("サブフォルダからも検索する(&S)")
        self.chk_first = QCheckBox("ファイル毎最初のみ検索(&G)")
        v2.addWidget(self.chk_sub, 0, 0, 1, 3)
        v2.addWidget(self.chk_first, 1, 0, 1, 3)
        kinds = QHBoxLayout()
        self.kind_checks = {}
        for k, label in (("text", "テキスト"), ("excel", "Excel"), ("word", "Word"),
                         ("ppt", "PowerPoint"), ("pdf", "PDF")):
            cb = QCheckBox(label)
            cb.setChecked(True)
            self.kind_checks[k] = cb
            kinds.addWidget(cb)
        kinds.addStretch(1)
        v2.addWidget(QLabel("種類:"), 2, 0)
        v2.addLayout(kinds, 2, 1, 1, 2)
        self.enc = QComboBox()
        for name, val in ENCODING_CHOICES:
            self.enc.addItem(name, val)
        el = QLabel("文字コードセット(&A)")
        el.setBuddy(self.enc)
        v2.addWidget(el, 3, 0)
        v2.addWidget(self.enc, 3, 1)

        g3 = QGroupBox("結果出力")
        v3 = QVBoxLayout(g3)
        self.rb_line = QRadioButton("該当行(&L)　…同じ行(セル)の複数ヒットは1件にまとめる")
        self.rb_part = QRadioButton("該当部分(&P)　…ヒットした箇所ごとに1件")
        self.rb_line.setChecked(True)
        bg = QButtonGroup(self)
        bg.addButton(self.rb_line)
        bg.addButton(self.rb_part)
        self.chk_cache = QCheckBox("抽出キャッシュを使う（Office・PDF の2回目以降が速い）")
        self.chk_cache.setChecked(True)
        v3.addWidget(self.rb_line)
        v3.addWidget(self.rb_part)
        v3.addWidget(self.chk_cache)
        v3.addStretch(1)
        opt_row.addWidget(g1)
        opt_row.addWidget(g2, 1)
        opt_row.addWidget(g3)
        grid.addLayout(opt_row, 5, 0, 1, 4)
        root.addWidget(box)

        # ---- 進捗
        prog_box = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress.setRange(0, 1000)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        self.progress.setFormat("")
        self.progress.setMinimumHeight(22)
        self.status = QLabel("条件を入れて「検索」を押してください。")
        self.status.setTextFormat(Qt.RichText)
        self.status.linkActivated.connect(lambda _: self.tabs.setCurrentIndex(2))
        prog_box.addWidget(self.progress, 2)
        prog_box.addWidget(self.status, 3)
        root.addLayout(prog_box)
        self.current_file = QLabel("")
        self.current_file.setStyleSheet("color:#666")
        root.addWidget(self.current_file)

        # ---- 結果
        self.hit_model = HitModel()
        self.proxy = SortProxy()
        self.proxy.setSourceModel(self.hit_model)
        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.table.setSortingEnabled(False)
        self.table.horizontalHeader().setSectionsClickable(True)
        self.table.horizontalHeader().sectionClicked.connect(self._sort_hits)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(24)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(False)
        self.table.setMouseTracking(True)
        self.table.setItemDelegateForColumn(HitModel.C_FILE, LinkDelegate(self.table))
        self.table.setItemDelegateForColumn(HitModel.C_CTX, HighlightDelegate(lambda: self.hit_model, self.table))
        self.table.clicked.connect(self._hit_clicked)
        self.table.doubleClicked.connect(lambda ix: self.open_hit(ix))
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._hit_menu)
        self.table.selectionModel().currentRowChanged.connect(self._show_preview)
        self.table.entered.connect(self._hover)
        hh = self.table.horizontalHeader()
        for c, w in enumerate([50, 200, 80, 120, 55, 50, 190, 440, 45, 260, 115]):
            hh.resizeSection(c, w)
        hh.setStretchLastSection(False)

        self.file_model = FileModel()
        self.file_table = self._make_file_table(self.file_model, [240, 80, 80, 80, 80, 115, 380, 200])
        self.err_model = FileModel(errors_mode=True)
        self.err_table = self._make_file_table(self.err_model, [240, 80, 70, 380, 380])

        self.tabs = QTabWidget()
        hit_page = QWidget()
        hv = QVBoxLayout(hit_page)
        hv.setContentsMargins(0, 4, 0, 0)
        filt = QHBoxLayout()
        self.filter = QLineEdit()
        self.filter.setPlaceholderText("結果をさらに絞り込む（ファイル名・場所・該当箇所に含まれる文字）")
        self.filter.setClearButtonEnabled(True)
        self.filter.textChanged.connect(self._filter_changed)
        filt.addWidget(QLabel("絞り込み:"))
        filt.addWidget(self.filter, 1)
        self.btn_export = QPushButton("Excelに出力(&E)")
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self.export_excel)
        filt.addWidget(self.btn_export)
        hv.addLayout(filt)
        hv.addWidget(self.table)
        legend = QLabel(LEGEND)
        legend.setWordWrap(True)
        legend.setStyleSheet("color:#444;font-size:11px;padding:2px")
        hv.addWidget(legend)
        self.tabs.addTab(hit_page, "検索結果")
        self.tabs.addTab(self.file_table, "ファイル別")
        self.tabs.addTab(self.err_table, "読めなかったファイル")

        self.preview = QTextBrowser()
        self.preview.setOpenLinks(False)
        self.preview.anchorClicked.connect(self._preview_link)
        self.preview.setPlaceholderText("結果の行を選ぶと、ここに該当箇所の前後が表示されます。")
        split = QSplitter(Qt.Vertical)
        split.addWidget(self.tabs)
        split.addWidget(self.preview)
        split.setStretchFactor(0, 4)
        split.setStretchFactor(1, 1)
        split.setSizes([520, 140])
        root.addWidget(split, 1)

        # ---- メニュー
        m_file = self.menuBar().addMenu("ファイル(&F)")
        act = QAction("Excelに出力(&E)…", self)
        act.setShortcut(QKeySequence("Ctrl+E"))
        act.triggered.connect(self.export_excel)
        m_file.addAction(act)
        m_file.addSeparator()
        act = QAction("終了(&X)", self)
        act.triggered.connect(self.close)
        m_file.addAction(act)
        m_set = self.menuBar().addMenu("設定(&S)")
        act = QAction("エディタ・キャッシュの設定(&O)…", self)
        act.triggered.connect(self.open_settings)
        m_set.addAction(act)
        m_help = self.menuBar().addMenu("ヘルプ(&H)")
        act = QAction("使い方(&H)", self)
        act.setShortcut(QKeySequence("F1"))
        act.triggered.connect(self.show_help)
        m_help.addAction(act)

        esc = QAction(self)
        esc.setShortcut(QKeySequence(Qt.Key_Escape))
        esc.triggered.connect(self.cancel_search)
        self.addAction(esc)

    def _make_file_table(self, model, widths):
        t = QTableView()
        proxy = QSortFilterProxyModel()
        proxy.setSourceModel(model)
        proxy.setSortRole(Qt.UserRole)
        t.setModel(proxy)
        t.setSortingEnabled(True)
        t.sortByColumn(-1, Qt.AscendingOrder)
        t.setSelectionBehavior(QAbstractItemView.SelectRows)
        t.setEditTriggers(QAbstractItemView.NoEditTriggers)
        t.verticalHeader().setVisible(False)
        t.verticalHeader().setDefaultSectionSize(24)
        t.setAlternatingRowColors(True)
        t.setMouseTracking(True)
        t.setItemDelegateForColumn(0, LinkDelegate(t))
        for c, w in enumerate(widths):
            t.horizontalHeader().resizeSection(c, w)

        def clicked(ix):
            if ix.column() == 0:
                p = model.path_at(proxy.mapToSource(ix).row())
                self._open_plain(p)

        t.clicked.connect(clicked)
        t.doubleClicked.connect(lambda ix: self._open_plain(model.path_at(proxy.mapToSource(ix).row())))
        t.entered.connect(lambda ix: t.viewport().setCursor(
            Qt.PointingHandCursor if ix.column() == 0 else Qt.ArrowCursor))
        t.setContextMenuPolicy(Qt.CustomContextMenu)

        def menu(pos):
            ix = t.indexAt(pos)
            if not ix.isValid():
                return
            p = model.path_at(proxy.mapToSource(ix).row())
            m = QMenu(t)
            m.addAction("ファイルを開く", lambda: self._open_plain(p))
            m.addAction("フォルダを開く（ファイルを選択）", lambda: opener.reveal_in_explorer(p))
            m.addAction("フルパスをコピー", lambda: QGuiApplication.clipboard().setText(p))
            m.exec(t.viewport().mapToGlobal(pos))

        t.customContextMenuRequested.connect(menu)
        return t

    # ------------------------------------------------------------------ 設定の保存・復元
    def _hist(self, combo: QComboBox, key: str, default=None):
        items = self.settings.value(key, []) or []
        if isinstance(items, str):
            items = [items]
        combo.addItems(items)
        if default is not None and not items:
            combo.setCurrentText(default)
        elif items:
            combo.setCurrentIndex(0)

    def _push_hist(self, combo: QComboBox, key: str):
        text = combo.currentText()
        items = [combo.itemText(i) for i in range(combo.count())]
        if text in items:
            items.remove(text)
        items.insert(0, text)
        items = items[:HIST_MAX]
        combo.blockSignals(True)
        combo.clear()
        combo.addItems(items)
        combo.setCurrentIndex(0)
        combo.blockSignals(False)
        self.settings.setValue(key, items)

    def _load_settings(self):
        s = self.settings
        self._hist(self.cond, "hist/cond")
        self.cond.setCurrentText("")
        self._hist(self.include, "hist/include", "*.*")
        self._hist(self.folder, "hist/folder", os.getcwd())
        self._hist(self.ex_files, "hist/ex_files", "")
        self._hist(self.ex_dirs, "hist/ex_dirs", ".git;.svn;node_modules")

        def b(key, default):
            v = s.value(key, default)
            return v in (True, "true", "1", 1)

        self.chk_word.setChecked(b("opt/word", False))
        self.chk_case.setChecked(b("opt/case", False))
        self.chk_regex.setChecked(b("opt/regex", False))
        self.chk_sub.setChecked(b("opt/sub", True))
        self.chk_first.setChecked(b("opt/first", False))
        self.chk_cache.setChecked(b("opt/cache", True))
        (self.rb_part if s.value("opt/mode", "line") == "part" else self.rb_line).setChecked(True)
        for k, cb in self.kind_checks.items():
            cb.setChecked(b(f"kind/{k}", True))
        i = self.enc.findData(s.value("opt/enc", "auto"))
        self.enc.setCurrentIndex(max(0, i))
        geo = s.value("win/geometry")
        if geo:
            self.restoreGeometry(geo)

    def _save_opts(self):
        s = self.settings
        s.setValue("opt/word", self.chk_word.isChecked())
        s.setValue("opt/case", self.chk_case.isChecked())
        s.setValue("opt/regex", self.chk_regex.isChecked())
        s.setValue("opt/sub", self.chk_sub.isChecked())
        s.setValue("opt/first", self.chk_first.isChecked())
        s.setValue("opt/cache", self.chk_cache.isChecked())
        s.setValue("opt/mode", "part" if self.rb_part.isChecked() else "line")
        s.setValue("opt/enc", self.enc.currentData())
        for k, cb in self.kind_checks.items():
            s.setValue(f"kind/{k}", cb.isChecked())

    def closeEvent(self, e):
        if self.thread and self.thread.isRunning():
            self.thread.cancel.set()
            self.thread.wait(3000)
        self.settings.setValue("win/geometry", self.saveGeometry())
        self._save_opts()
        super().closeEvent(e)

    # ------------------------------------------------------------------ フォルダ操作
    def browse_folder(self):
        cur = self.folder.currentText().split(";")[0].strip()
        d = QFileDialog.getExistingDirectory(self, "検索するフォルダ", cur if os.path.isdir(cur) else "")
        if d:
            self.folder.setCurrentText(os.path.normpath(d))

    def parent_folder(self):
        cur = self.folder.currentText().split(";")[0].strip()
        if cur:
            self.folder.setCurrentText(os.path.dirname(os.path.normpath(cur)) or cur)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        paths = [u.toLocalFile() for u in e.mimeData().urls() if u.isLocalFile()]
        dirs = [p if os.path.isdir(p) else os.path.dirname(p) for p in paths]
        if dirs:
            self.folder.setCurrentText(";".join(os.path.normpath(d) for d in dict.fromkeys(dirs)))

    # ------------------------------------------------------------------ 検索
    def start_search(self):
        if self.thread and self.thread.isRunning():
            return
        pattern = self.cond.currentText()
        if not pattern:
            QMessageBox.information(self, "検索", "条件（検索する文字列）を入力してください。")
            self.cond.setFocus()
            return
        folders = [os.path.expandvars(os.path.expanduser(p.strip().strip('"')))
                   for p in self.folder.currentText().split(";") if p.strip()]
        bad = [p for p in folders if not os.path.isdir(p)]
        if not folders or bad:
            QMessageBox.warning(self, "検索", "フォルダが見つかりません:\n" + "\n".join(bad or ["（未入力）"]))
            self.folder.setFocus()
            return
        kinds = [k for k, cb in self.kind_checks.items() if cb.isChecked()]
        if not kinds:
            QMessageBox.information(self, "検索", "対象の種類を1つ以上選んでください。")
            return
        opts = dict(
            pattern=pattern, regex=self.chk_regex.isChecked(), word=self.chk_word.isChecked(),
            case=self.chk_case.isChecked(), encoding=self.enc.currentData(),
            mode="part" if self.rb_part.isChecked() else "line", first_only=self.chk_first.isChecked(),
            use_cache=self.chk_cache.isChecked() and self.cache is not None,
            cache_path=default_cache_path() if self.cache else None,
            folders=folders, include=self.include.currentText() or "*.*",
            exclude_files=self.ex_files.currentText(), exclude_dirs=self.ex_dirs.currentText(),
            recursive=self.chk_sub.isChecked(), kinds=kinds,
        )
        if opts["regex"]:
            try:
                re.compile(pattern)
            except re.error as e:
                QMessageBox.warning(self, "正規表現", f"正規表現に誤りがあります:\n{e}")
                return
        for combo, key in ((self.cond, "hist/cond"), (self.include, "hist/include"),
                           (self.folder, "hist/folder"), (self.ex_files, "hist/ex_files"),
                           (self.ex_dirs, "hist/ex_dirs")):
            self._push_hist(combo, key)
        self._save_opts()
        self.last_opts = opts
        self.started_at = dt.datetime.now()

        self.hit_model.reset([])
        self.file_model.reset()
        self.err_model.reset()
        self.filter.clear()
        self.preview.clear()
        self.files = []
        self.summary = None
        self._update_tabs(0, 0, 0)
        self.btn_search.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.btn_export.setEnabled(False)
        self.progress.setRange(0, 0)
        self.progress.setFormat("")
        self.status.setText("対象ファイルを数えています…")
        self.current_file.setText("")

        self.thread = SearchThread(self.engine, opts)
        self.thread.sig_enum.connect(self._on_enum)
        self.thread.sig_started.connect(self._on_started)
        self.thread.sig_progress.connect(self._on_progress)
        self.thread.sig_results.connect(self._on_results)
        self.thread.sig_finished.connect(self._on_finished)
        self.thread.sig_error.connect(self._on_error)
        self.thread.start()

    def cancel_search(self):
        if self.thread and self.thread.isRunning():
            self.thread.cancel.set()
            self.status.setText("キャンセルしています…")

    def _on_enum(self, n):
        self.status.setText(f"対象ファイルを数えています… {n:,} 件")

    def _on_started(self, files, total_bytes):
        self.files = files
        self.hit_model.files = files
        self.progress.setRange(0, 1000)
        self.progress.setValue(0)
        self.status.setText(f"検索中… 0 / {len(files):,} ファイル（合計 {fmt_size(total_bytes)}）")

    def _on_progress(self, done, total, ratio, eta, cur, hits):
        self.progress.setValue(int(ratio * 1000))
        self.progress.setFormat(f"{ratio * 100:.0f}%")
        el = (dt.datetime.now() - self.started_at).total_seconds()
        self.status.setText(
            f"<b>残り {fmt_secs(eta)}</b>　｜　{done:,} / {total:,} ファイル　｜　"
            f"ヒット {hits:,} 件　｜　経過 {fmt_elapsed(el)}")
        if cur:
            self.current_file.setText("処理中: " + cur)

    def _on_results(self, files, hits):
        self.hit_model.append(hits)
        self.file_model.append([f for f in files if f.hits > 0])
        self.err_model.append([f for f in files if f.status == "error" or
                               (f.status == "skip" and f.kind != "text")])
        self._update_tabs(len(self.hit_model.hits), len(self.file_model.rows), len(self.err_model.rows))

    def _update_tabs(self, hits, files, errs):
        self.tabs.setTabText(0, f"検索結果 ({hits:,})")
        self.tabs.setTabText(1, f"ファイル別 ({files:,})")
        self.tabs.setTabText(2, f"読めなかったファイル ({errs:,})" if errs else "読めなかったファイル")

    def _on_finished(self, files, summ):
        self.summary = summ
        self.files = files
        if summ.enum_errors:
            self.err_model.append(list(summ.enum_errors))
        self.hit_model.finalize_order()
        self._update_tabs(len(self.hit_model.hits), len(self.file_model.rows), len(self.err_model.rows))
        self.btn_search.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.btn_export.setEnabled(bool(self.hit_model.hits))
        self.progress.setRange(0, 1000)
        self.progress.setValue(1000 if not summ.cancelled else self.progress.value())
        self.progress.setFormat("完了" if not summ.cancelled else "中断")
        self.current_file.setText("")
        self.status.setText(self._summary_html(summ))
        self.statusBar().showMessage(self._summary_text(summ), 0)
        if self.hit_model.hits:
            self.table.selectRow(0)

    def _summary_text(self, s) -> str:
        head = "キャンセルしました" if s.cancelled else ("ヒット件数が上限に達したため打ち切りました" if s.hit_limit else "完了")
        extra = []
        if s.cached:
            extra.append(f"キャッシュ利用 {s.cached:,}")
        if s.errors or s.enum_errors:
            extra.append(f"読めなかった {s.errors + len(s.enum_errors):,}")
        if s.skipped:
            extra.append(f"バイナリ等でスキップ {s.skipped:,}")
        return (f"{head}: {s.files_done:,} / {s.files_total:,} ファイル中 {s.files_hit:,} ファイルで "
                f"{s.hits:,} 件ヒット（{s.elapsed:.1f}秒" + ("、" + "、".join(extra) if extra else "") + "）")

    def _summary_html(self, s) -> str:
        t = html.escape(self._summary_text(s))
        if s.errors or s.enum_errors:
            t += "　<a href='errors'>読めなかったファイルを見る</a>"
        return t

    def _on_error(self, msg):
        self.btn_search.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.progress.setRange(0, 1000)
        self.progress.setValue(0)
        self.progress.setFormat("")
        self.status.setText(html.escape(msg))
        QMessageBox.warning(self, "検索", msg)

    # ------------------------------------------------------------------ 結果操作
    def _sort_hits(self, col):
        order = Qt.AscendingOrder
        if self.proxy.sortColumn() == col and self.proxy.sortOrder() == Qt.AscendingOrder:
            order = Qt.DescendingOrder
        self.proxy.sort(col, order)
        self.table.horizontalHeader().setSortIndicatorShown(True)
        self.table.horizontalHeader().setSortIndicator(col, order)

    def _filter_changed(self, text):
        self.proxy.set_text(text)
        n = self.proxy.rowCount()
        total = len(self.hit_model.hits)
        self.tabs.setTabText(0, f"検索結果 ({n:,} / {total:,})" if text else f"検索結果 ({total:,})")

    def _hit_at(self, proxy_index):
        if not proxy_index.isValid():
            return None
        return self.hit_model.hits[self.proxy.mapToSource(proxy_index).row()]

    def _hover(self, ix):
        self.table.viewport().setCursor(Qt.PointingHandCursor if ix.column() == HitModel.C_FILE else Qt.ArrowCursor)

    def _hit_clicked(self, ix):
        if ix.column() == HitModel.C_FILE:
            self.open_hit(ix)

    def open_hit(self, ix, plain=False):
        h = self._hit_at(ix)
        if not h:
            return
        f = self.files[h[0]]
        if not os.path.exists(f.path):
            QMessageBox.warning(self, "開く", f"ファイルが見つかりません:\n{f.path}")
            return
        if plain:
            self._open_plain(f.path)
            return
        self.statusBar().showMessage(f"開いています: {f.path}", 4000)
        opener.open_at(f.path, h[1], h[2], h[3], h[4], h[5], h[6], self.editor or "",
                       on_error=lambda m: self.statusBar().showMessage(m, 8000))

    def _open_plain(self, path):
        try:
            opener.open_default(path)
        except Exception as e:
            QMessageBox.warning(self, "開く", f"開けませんでした:\n{e}")

    def _hit_menu(self, pos):
        ix = self.table.indexAt(pos)
        h = self._hit_at(ix)
        if not h:
            return
        f = self.files[h[0]]
        m = QMenu(self)
        m.addAction("該当箇所を開く", lambda: self.open_hit(ix))
        m.addAction("ファイルを開く（先頭から）", lambda: self.open_hit(ix, plain=True))
        m.addAction("フォルダを開く（ファイルを選択）", lambda: opener.reveal_in_explorer(f.path))
        m.addSeparator()
        m.addAction("フルパスをコピー", lambda: QGuiApplication.clipboard().setText(f.path))
        m.addAction("選択行をコピー（タブ区切り）", self.copy_selection)
        m.exec(self.table.viewport().mapToGlobal(pos))

    def copy_selection(self):
        rows = sorted({ix.row() for ix in self.table.selectionModel().selectedRows()})
        out = ["\t".join(HitModel.COLS)]
        for r in rows:
            vals = []
            for c in range(len(HitModel.COLS)):
                v = self.proxy.index(r, c).data(Qt.DisplayRole)
                vals.append("" if v is None else str(v))
            out.append("\t".join(vals))
        QGuiApplication.clipboard().setText("\n".join(out))

    def keyPressEvent(self, e):
        if e.matches(QKeySequence.Copy) and self.table.hasFocus():
            self.copy_selection()
            return
        if e.key() in (Qt.Key_Return, Qt.Key_Enter) and self.table.hasFocus():
            self.open_hit(self.table.currentIndex())
            return
        super().keyPressEvent(e)

    def _show_preview(self, cur, _prev):
        h = self._hit_at(cur)
        if not h:
            return
        f = self.files[h[0]]
        kind, place, line, col, label = h[1], h[2], h[3], h[4], h[5]
        loc = []
        if place:
            loc.append(html.escape(place))
        if line is not None:
            loc.append(f"{'段落' if kind == 'word' else '行'} {line}")
        if col is not None:
            loc.append(f"{'列' if kind == 'excel' else '桁'} {col}")
        if label:
            loc.append(html.escape(label))
        head = (f"<a href='open'><b>{html.escape(os.path.basename(f.path))}</b></a>　"
                f"<span style='color:#666'>{html.escape(KIND_LABEL.get(kind, kind))}　{' ／ '.join(loc)}</span><br>"
                f"<span style='color:#888'>{html.escape(f.path)}</span>")
        body = ""
        if kind == "text" and line:
            body = self._text_context(f, line, h)
        if not body:
            ctx = h[7]
            body = ("<pre style='margin:4px 0'>" + html.escape(ctx[:h[8]]) +
                    "<span style='background:#ffe066;font-weight:bold'>" + html.escape(ctx[h[8]:h[9]]) +
                    "</span>" + html.escape(ctx[h[9]:]) + "</pre>")
        self.preview.setHtml(head + body)
        self._preview_ix = cur

    def _text_context(self, f, line, h, before=3, after=3) -> str:
        try:
            if f.size > 30 * 1024 * 1024:
                return ""
            with open(f.path, "rb") as fp:
                data = fp.read()
            enc = f.encoding if f.encoding and f.encoding != "ascii" else "auto"
            text, _ = decode_text(data, enc if enc != "auto" else "auto")
        except Exception:
            return ""
        lines = text.splitlines()
        s, e = max(1, line - before), min(len(lines), line + after)
        out = ["<pre style='margin:4px 0'>"]
        for n in range(s, e + 1):
            t = lines[n - 1].replace("\t", "    ") if n != line else lines[n - 1]
            num = f"<span style='color:#999'>{n:>6}: </span>"
            if n == line:
                c0 = (h[4] or 1) - 1
                ln = len(h[6])
                seg = (html.escape(t[:c0]) + "<span style='background:#ffe066;font-weight:bold'>" +
                       html.escape(t[c0:c0 + ln]) + "</span>" + html.escape(t[c0 + ln:]))
                out.append(f"<div style='background:#fff7d6'>{num}{seg}</div>")
            else:
                out.append(f"<div>{num}{html.escape(t)}</div>")
        out.append("</pre>")
        return "".join(out)

    def _preview_link(self, url):
        if getattr(self, "_preview_ix", None) is not None:
            self.open_hit(self._preview_ix)

    # ------------------------------------------------------------------ Excel出力
    def export_excel(self):
        if not self.hit_model.hits:
            QMessageBox.information(self, "Excel出力", "出力する検索結果がありません。")
            return
        if self.thread and self.thread.isRunning():
            QMessageBox.information(self, "Excel出力", "検索が終わってから出力してください。")
            return
        default_dir = self.settings.value("export/dir", os.path.join(os.path.expanduser("~"), "Documents"))
        if not os.path.isdir(default_dir):
            default_dir = os.path.expanduser("~")
        safe = re.sub(r'[\\/:*?"<>|]', "_", self.last_opts["pattern"])[:30] if self.last_opts else "grep"
        name = f"grep_{safe}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Excelに出力", os.path.join(default_dir, name),
                                              "Excel ブック (*.xlsx)")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        self.settings.setValue("export/dir", os.path.dirname(path))
        # 画面の並び順・絞り込みのとおりに出力する
        hits = [self.hit_model.hits[self.proxy.mapToSource(self.proxy.index(r, 0)).row()]
                for r in range(self.proxy.rowCount())]
        dlg = QProgressDialog("Excel ファイルを作成しています…", None, 0, max(1, len(hits)), self)
        dlg.setWindowTitle("Excel出力")
        dlg.setWindowModality(Qt.WindowModal)
        dlg.setMinimumDuration(300)
        self._export_thread = ExportThread(path, hits, self.files, self._conditions(), [self._summary_text(self.summary)]
                                           if self.summary else [])
        self._export_thread.sig_progress.connect(lambda i, n: dlg.setValue(i))

        def done(err):
            dlg.close()
            if err:
                QMessageBox.warning(self, "Excel出力", err)
                return
            if QMessageBox.question(self, "Excel出力", f"出力しました。\n{path}\n\n今すぐ開きますか？") == QMessageBox.Yes:
                self._open_plain(path)

        self._export_thread.sig_done.connect(done)
        self._export_thread.start()

    def _conditions(self):
        o = self.last_opts or {}
        yn = lambda v: "する" if v else "しない"
        return [
            ("検索日時", self.started_at.strftime("%Y/%m/%d %H:%M:%S") if hasattr(self, "started_at") else ""),
            ("条件", o.get("pattern", "")),
            ("ファイル", o.get("include", "")),
            ("フォルダ", "; ".join(o.get("folders", []))),
            ("除外ファイル", o.get("exclude_files", "")),
            ("除外フォルダ", o.get("exclude_dirs", "")),
            ("単語単位", yn(o.get("word"))),
            ("大文字小文字を区別", yn(o.get("case"))),
            ("正規表現", yn(o.get("regex"))),
            ("サブフォルダ", yn(o.get("recursive"))),
            ("ファイル毎最初のみ", yn(o.get("first_only"))),
            ("対象の種類", "、".join(KIND_LABEL[k] for k in o.get("kinds", []))),
            ("文字コードセット", self.enc.currentText()),
            ("結果出力", "該当部分" if o.get("mode") == "part" else "該当行"),
            ("絞り込み", self.filter.text()),
        ]

    # ------------------------------------------------------------------ その他
    def open_settings(self):
        d = SettingsDialog(self, self.editor or "", self.cache)
        if d.exec() == QDialog.Accepted:
            self.editor = d.cmd.text().strip()
            self.settings.setValue("editor", self.editor)

    def show_help(self):
        d = QDialog(self)
        d.setWindowTitle("使い方")
        d.resize(720, 640)
        v = QVBoxLayout(d)
        tb = QTextBrowser()
        tb.setHtml(HELP_HTML)
        v.addWidget(tb)
        bb = QDialogButtonBox(QDialogButtonBox.Close)
        bb.rejected.connect(d.reject)
        v.addWidget(bb)
        d.exec()


def run(argv=None):
    argv = argv if argv is not None else sys.argv
    app = QApplication(argv)
    app.setApplicationName(APP)
    if sys.platform.startswith("win"):
        app.setFont(QFont("Meiryo UI", 9))
    app.setStyle("Fusion")
    try:  # 配色はライトに固定 (強調表示や凡例の色が見やすいように)
        app.styleHints().setColorScheme(Qt.ColorScheme.Light)
    except Exception:
        pass
    engine = SearchEngine()
    engine.warmup()
    folder = argv[1] if len(argv) > 1 and os.path.isdir(argv[1]) else None
    w = MainWindow(engine, folder)
    w.show()
    rc = app.exec()
    engine.shutdown()
    return rc
