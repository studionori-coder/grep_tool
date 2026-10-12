"""テストデータを実際に検索し、期待結果 (make_testdata.EXPECTED) と照合する。

  python tests/check_testdata.py <テストデータのフォルダ> [結果を書き出すファイル]
"""
from __future__ import annotations

import os
import sys
import threading
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
sys.path.insert(0, HERE)

from greptool.engine import SearchEngine  # noqa: E402
from make_testdata import EXPECTED, EXPECTED_ERRORS, EXPECTED_NOT_FOUND  # noqa: E402


class CB:
    def __init__(self):
        self.hits = []

    def enumerating(self, n):
        pass

    def started(self, files, total):
        pass

    def progress(self, *a):
        pass

    def results(self, files, hits):
        self.hits.extend(hits)


def search(engine, root, pattern, use_cache=False, cache_path=None):
    cb = CB()
    opts = dict(pattern=pattern, regex=False, word=False, case=True, encoding="auto", mode="part",
                first_only=False, use_cache=use_cache, cache_path=cache_path, folders=[root], include="*.*",
                exclude_files="", exclude_dirs=".git;.svn;node_modules", recursive=True,
                kinds=["text", "excel", "word", "ppt", "pdf"])
    files, summ = engine.run(opts, cb, threading.Event())
    return files, summ, cb.hits


def fmt(v):
    return "－" if v is None else str(v)


def main():
    root = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "testdata")
    out_path = sys.argv[2] if len(sys.argv) > 2 else None
    lines = []
    p = lambda s="": (print(s), lines.append(s))
    engine = SearchEngine()
    ok = ng = 0

    p("■ 期待結果との照合（検索語ごとに検索し、ファイル・場所・行・列・位置を確認）")
    by_kw: dict[str, list] = {}
    for e in EXPECTED:
        by_kw.setdefault(e[1], []).append(e)
    for kw, exps in by_kw.items():
        files, summ, hits = search(engine, root, kw)
        for rel, _, kind, place, line, col, label in exps:
            path = os.path.join(root, rel)
            cands = [h for h in hits if os.path.normcase(files[h[0]].path) == os.path.normcase(path)]
            match = None
            for h in cands:
                if h[1] != kind:
                    continue
                if place != "?" and h[2] != place:
                    continue
                if line != "?" and h[3] != line:
                    continue
                if col != "?" and h[4] != col:
                    continue
                if label and label not in (h[5] or ""):
                    continue
                match = h
                break
            want = f"{place} 行{fmt(line)} 列{fmt(col)} {label}".strip()
            if match:
                ok += 1
                p(f"  OK  {kw:<22} {rel}  → {match[2]} 行{fmt(match[3])} 列{fmt(match[4])} {match[5]}")
            else:
                ng += 1
                got = "; ".join(f"{h[2]} 行{fmt(h[3])} 列{fmt(h[4])} {h[5]}" for h in cands) or "ヒットなし"
                st = next((f for f in files if os.path.normcase(f.path) == os.path.normcase(path)), None)
                extra = f" [状態: {st.status} {st.message}]" if st and st.status in ("error", "skip") else ""
                p(f"  NG  {kw:<22} {rel}\n        期待: {want}\n        実際: {got}{extra}")

    p()
    p("■ 読めなかったファイル（理由の表示）")
    files, summ, hits = search(engine, root, "ORD_STATUS")
    errs = {os.path.relpath(f.path, root).replace(os.sep, "/"): f for f in files if f.status in ("error", "skip")}
    for rel, reason in EXPECTED_ERRORS:
        f = errs.get(rel)
        if f and reason in f.message:
            ok += 1
            p(f"  OK  {rel}  → {f.message}")
        else:
            ng += 1
            p(f"  NG  {rel}  期待: 「{reason}」を含む理由  実際: {f.message if f else '一覧に無い'}")
    p()
    p("■ 検索対象外になるべきファイル（除外フォルダ・Office一時ファイル・バイナリ）")
    listed = {os.path.relpath(f.path, root).replace(os.sep, "/") for f in files}
    for rel in EXPECTED_NOT_FOUND:
        if rel not in listed:
            ok += 1
            p(f"  OK  {rel}  → 対象外")
        else:
            ng += 1
            p(f"  NG  {rel}  → 検索対象になっている")

    p()
    p("■ 速度（全ファイル、抽出キャッシュあり）")
    cache = os.path.join(HERE, "_check_cache.sqlite")
    for suffix in ("", "-wal", "-shm"):
        if os.path.exists(cache + suffix):
            os.remove(cache + suffix)
    from greptool.cache import ExtractCache
    ExtractCache(cache).close()
    for i in range(2):
        t = time.time()
        files, summ, hits = search(engine, root, "受注番号", True, cache)
        p(f"  {i + 1}回目: {summ.files_total} ファイル, ヒット {summ.hits} 件, "
          f"読めなかった {summ.errors} 件, キャッシュ利用 {summ.cached} 件, {time.time() - t:.2f} 秒")
    for suffix in ("", "-wal", "-shm"):
        if os.path.exists(cache + suffix):
            os.remove(cache + suffix)
    engine.shutdown()
    p()
    p(f"結果: OK {ok} 件 / NG {ng} 件")
    if out_path:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
    return 0 if ng == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
