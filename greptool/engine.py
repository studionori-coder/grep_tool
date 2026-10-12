"""検索の段取り (ファイル列挙 → 並列検索 → 進捗・残り時間の計算)。Qt に依存しない。"""
from __future__ import annotations

import fnmatch
import multiprocessing
import os
import re
import threading
import time
from collections import deque
from concurrent.futures import FIRST_COMPLETED, ProcessPoolExecutor, wait
from concurrent.futures.process import BrokenProcessPool
from dataclasses import dataclass, field

from . import searcher
from .cache import ExtractCache
from .extractors import KIND_TEXT, kind_of

MAX_TOTAL_HITS = 300_000
PER_FILE_COST = 32 * 1024  # 1ファイルあたりの固定コスト (バイト換算) — 残り時間の見積もり用


@dataclass
class FileInfo:
    index: int
    path: str
    kind: str
    size: int
    mtime: float
    status: str = "pending"   # pending / ok / cached / error / skip
    message: str = ""
    hits: int = 0
    encoding: str = ""
    truncated: bool = False


@dataclass
class Summary:
    files_total: int = 0
    files_done: int = 0
    files_hit: int = 0
    hits: int = 0
    errors: int = 0
    skipped: int = 0
    cached: int = 0
    elapsed: float = 0.0
    cancelled: bool = False
    hit_limit: bool = False
    enum_errors: list = field(default_factory=list)


def split_patterns(s: str) -> list[str]:
    pats = [p.strip() for p in re.split(r"[;,\s]+", s or "") if p.strip()]
    return ["*" if p in ("*.*", "**") else p.lower() for p in pats]


def _match_any(name: str, pats: list[str]) -> bool:
    n = name.lower()
    return any(fnmatch.fnmatchcase(n, p) for p in pats)


def enumerate_files(roots, include, exclude_files, exclude_dirs, recursive, kinds, cancel, on_count=None):
    """条件に合うファイルを列挙する。戻り値: (FileInfo のリスト, 列挙時のエラー)"""
    inc = split_patterns(include) or ["*"]
    exf = split_patterns(exclude_files)
    exd = split_patterns(exclude_dirs)
    files: list[FileInfo] = []
    errors: list[tuple[str, str]] = []
    last = time.monotonic()
    stack = [r for r in roots]
    while stack:
        if cancel.is_set():
            break
        d = stack.pop()
        try:
            with os.scandir(d) as it:
                entries = sorted(it, key=lambda e: e.name.lower())
        except PermissionError:
            errors.append((d, "フォルダへのアクセス権がありません"))
            continue
        except OSError as e:
            errors.append((d, f"フォルダを開けません: {e.strerror or e}"))
            continue
        subdirs = []
        for e in entries:
            try:
                if e.is_dir(follow_symlinks=False):
                    if recursive and not (exd and _match_any(e.name, exd)):
                        subdirs.append(e.path)
                    continue
                if not e.is_file():
                    continue
            except OSError:
                continue
            name = e.name
            if name.startswith("~$"):  # Office の一時ファイル
                continue
            if not _match_any(name, inc) or (exf and _match_any(name, exf)):
                continue
            k = kind_of(name)
            if k is None or k not in kinds:
                continue
            try:
                st = e.stat()
            except OSError as ex:
                errors.append((e.path, f"情報を取得できません: {ex.strerror or ex}"))
                continue
            files.append(FileInfo(len(files), e.path, k, st.st_size, st.st_mtime))
        # 名前順に処理されるよう逆順で積む
        stack.extend(reversed(subdirs))
        now = time.monotonic()
        if on_count and now - last > 0.1:
            on_count(len(files))
            last = now
    if on_count:
        on_count(len(files))
    return files, errors


def _make_batches(files: list[FileInfo]):
    """小さいファイルはまとめ、大きいファイルは単独で1タスクにする"""
    batches = []
    cur, cur_cost, cur_kind = [], 0, None
    for f in files:
        heavy = f.kind != KIND_TEXT
        limit_n = 8 if heavy else 64
        limit_bytes = 2 * 1024 * 1024 if heavy else 8 * 1024 * 1024
        if cur and (len(cur) >= limit_n or cur_cost + f.size > limit_bytes or cur_kind != heavy):
            batches.append(cur)
            cur, cur_cost = [], 0
        cur.append(f)
        cur_cost += f.size
        cur_kind = heavy
    if cur:
        batches.append(cur)
    return batches


class Progress:
    """処理済みコストの推移から残り時間を見積もる"""

    def __init__(self, total_cost: int):
        self.total = max(1, total_cost)
        self.done = 0
        self.t0 = time.monotonic()
        self.hist = deque()  # (時刻, 処理済みコスト)

    def add(self, cost: int):
        self.done += cost
        now = time.monotonic()
        self.hist.append((now, self.done))
        while len(self.hist) > 2 and now - self.hist[0][0] > 8.0:
            self.hist.popleft()

    def eta(self) -> float | None:
        now = time.monotonic()
        elapsed = now - self.t0
        if self.done <= 0 or elapsed < 0.5:
            return None
        # 直近8秒の速度と全体平均の速度を混ぜる (急な変化に追従しつつ安定させる)
        overall = self.done / elapsed
        if len(self.hist) >= 2 and self.hist[-1][0] - self.hist[0][0] > 0.5:
            recent = (self.hist[-1][1] - self.hist[0][1]) / (self.hist[-1][0] - self.hist[0][0])
            rate = 0.6 * recent + 0.4 * overall
        else:
            rate = overall
        if rate <= 0:
            return None
        return max(0.0, (self.total - self.done) / rate)

    @property
    def ratio(self) -> float:
        return min(1.0, self.done / self.total)


class SearchEngine:
    def __init__(self, workers: int | None = None):
        n = workers or max(1, (os.cpu_count() or 2) - 1)
        self.workers = min(n, 12)
        self._pool: ProcessPoolExecutor | None = None
        self._lock = threading.Lock()

    def pool(self) -> ProcessPoolExecutor:
        with self._lock:
            if self._pool is None:
                ctx = multiprocessing.get_context("spawn")
                self._pool = ProcessPoolExecutor(max_workers=self.workers, mp_context=ctx)
            return self._pool

    def warmup(self):
        """アプリ起動直後にワーカーを立ち上げておく (初回検索の待ち時間を減らす)"""
        try:
            p = self.pool()
            for _ in range(self.workers):
                p.submit(searcher.warmup)
        except Exception:
            pass

    def _reset_pool(self):
        with self._lock:
            if self._pool is not None:
                try:
                    self._pool.shutdown(wait=False, cancel_futures=True)
                except Exception:
                    pass
            self._pool = None

    def shutdown(self):
        with self._lock:
            if self._pool is not None:
                self._pool.shutdown(wait=False, cancel_futures=True)
                self._pool = None

    def run(self, opts: dict, cb, cancel: threading.Event) -> tuple[list[FileInfo], Summary]:
        """cb は次のメソッドを持つオブジェクト:
            enumerating(count) / started(files, total_bytes) /
            progress(done_files, total_files, ratio, eta, current_path, hits) /
            results(updated_fileinfos, hits)
        """
        t0 = time.monotonic()
        summ = Summary()
        searcher.build_regex(opts)  # 正規表現の誤りはここで例外にする
        roots = [r for r in opts["folders"] if r]
        files, enum_errors = enumerate_files(
            roots, opts["include"], opts["exclude_files"], opts["exclude_dirs"],
            opts["recursive"], set(opts["kinds"]), cancel, cb.enumerating)
        summ.enum_errors = enum_errors
        summ.files_total = len(files)
        total_bytes = sum(f.size for f in files)
        cb.started(files, total_bytes)
        prog = Progress(total_bytes + PER_FILE_COST * len(files))

        cache = None
        if opts.get("use_cache") and opts.get("cache_path"):
            try:
                cache = ExtractCache(opts["cache_path"])
            except Exception:
                cache = None
                opts = dict(opts, use_cache=False)
        to_store = []

        wopts = {k: opts.get(k) for k in ("pattern", "regex", "word", "case", "encoding", "mode",
                                          "first_only", "use_cache", "cache_path")}
        batches = deque(_make_batches(files))
        inflight = {}
        max_inflight = self.workers * 3
        last_emit = 0.0
        pend_hits, pend_files = [], []
        current = ""

        def flush(force=False):
            nonlocal last_emit, pend_hits, pend_files
            now = time.monotonic()
            if force or now - last_emit > 0.15:
                if pend_hits or pend_files:
                    cb.results(pend_files, pend_hits)
                    pend_hits, pend_files = [], []
                cb.progress(summ.files_done, summ.files_total, prog.ratio, prog.eta(), current, summ.hits)
                last_emit = now

        try:
            while (batches or inflight) and not cancel.is_set():
                while batches and len(inflight) < max_inflight:
                    b = batches.popleft()
                    args = [(f.index, f.path, f.kind, f.size, f.mtime) for f in b]
                    try:
                        fut = self.pool().submit(searcher.process_batch, args, wopts)
                    except (BrokenProcessPool, RuntimeError):
                        self._reset_pool()
                        fut = self.pool().submit(searcher.process_batch, args, wopts)
                    inflight[fut] = b
                done, _ = wait(list(inflight), timeout=0.1, return_when=FIRST_COMPLETED)
                for fut in done:
                    b = inflight.pop(fut)
                    try:
                        res = fut.result()
                    except BrokenProcessPool:
                        self._reset_pool()
                        res = [(f.index, "error", "読み取り中に処理が異常終了しました (ファイルが壊れている可能性)",
                                [], "", None, False) for f in b]
                    except Exception as e:
                        res = [(f.index, "error", f"読み取りエラー: {e}", [], "", None, False) for f in b]
                    for fidx, status, msg, hits, enc, packed, trunc in res:
                        fi = files[fidx]
                        fi.status, fi.message, fi.encoding, fi.truncated = status, msg, enc, trunc
                        fi.hits = sum(h[10] for h in hits)
                        summ.files_done += 1
                        if status == "error":
                            summ.errors += 1
                        elif status == "skip":
                            summ.skipped += 1
                        elif status == "cached":
                            summ.cached += 1
                        if hits:
                            summ.files_hit += 1
                            summ.hits += len(hits)
                            pend_hits.extend(hits)
                        pend_files.append(fi)
                        if packed is not None and cache is not None:
                            to_store.append((fi.path, fi.mtime, fi.size, packed))
                        prog.add(fi.size + PER_FILE_COST)
                        current = fi.path
                    if len(to_store) >= 50:
                        cache.put_many(to_store)
                        to_store.clear()
                    if summ.hits >= MAX_TOTAL_HITS:
                        summ.hit_limit = True
                        cancel.set()
                flush()
        finally:
            if cancel.is_set():
                for fut in inflight:
                    fut.cancel()
            if cache is not None:
                try:
                    cache.put_many(to_store)
                except Exception:
                    pass
                cache.close()
        summ.cancelled = cancel.is_set() and not summ.hit_limit
        summ.elapsed = time.monotonic() - t0
        flush(force=True)
        return files, summ
