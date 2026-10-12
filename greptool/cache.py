"""抽出キャッシュ (簡易インデックス)。

Excel / Word / PowerPoint / PDF から取り出したテキストを、ファイルのパス・更新日時・サイズを
キーにして SQLite に保存する。2回目以降の検索ではファイルを開かずにキャッシュを検索するので
大幅に速くなる。ファイルが更新されると自動で読み直す。
"""
from __future__ import annotations

import json
import os
import sqlite3
import zlib

CACHE_VERSION = 4  # 抽出ロジックを変えたら上げる (古いキャッシュを無効化)


def default_cache_path() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), ".cache")
    d = os.path.join(base, "GrepTool")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, "extract_cache.sqlite")


def pack_units(units) -> bytes:
    return zlib.compress(json.dumps(units, ensure_ascii=False, separators=(",", ":")).encode("utf-8"), 3)


def unpack_units(blob: bytes):
    return [tuple(u) for u in json.loads(zlib.decompress(blob).decode("utf-8"))]


class ExtractCache:
    def __init__(self, path: str, readonly: bool = False):
        self.path = path
        if readonly:
            # 書き込みはしないが、WAL の共有メモリを扱えるよう通常モードで開く
            self.conn = sqlite3.connect(path, timeout=10, check_same_thread=False)
        else:
            self.conn = sqlite3.connect(path, timeout=30, check_same_thread=False)
            self.conn.execute("PRAGMA journal_mode=WAL")
            self.conn.execute("PRAGMA synchronous=NORMAL")
            self.conn.execute(
                "CREATE TABLE IF NOT EXISTS files (path TEXT PRIMARY KEY, mtime REAL, size INTEGER,"
                " ver INTEGER, data BLOB)")
            self.conn.commit()

    def get(self, path: str, mtime: float, size: int):
        try:
            row = self.conn.execute(
                "SELECT mtime, size, ver, data FROM files WHERE path=?", (path,)).fetchone()
        except sqlite3.Error:
            return None
        if row and abs(row[0] - mtime) < 1e-3 and row[1] == size and row[2] == CACHE_VERSION:
            try:
                return unpack_units(row[3])
            except Exception:
                return None
        return None

    def put_many(self, items):
        """items: [(path, mtime, size, packed_blob)]"""
        if not items:
            return
        self.conn.executemany(
            "INSERT OR REPLACE INTO files (path, mtime, size, ver, data) VALUES (?,?,?,?,?)",
            [(p, m, s, CACHE_VERSION, b) for p, m, s, b in items])
        self.conn.commit()

    def stats(self) -> tuple[int, int]:
        n = self.conn.execute("SELECT COUNT(*) FROM files").fetchone()[0]
        size = os.path.getsize(self.path) if os.path.exists(self.path) else 0
        return n, size

    def clear(self):
        self.conn.execute("DELETE FROM files")
        self.conn.commit()
        self.conn.execute("VACUUM")

    def close(self):
        self.conn.close()
