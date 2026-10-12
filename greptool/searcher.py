"""検索ワーカー (別プロセスで実行される)。Qt に依存しない。

ヒット (Hit) は次のタプル:
    (file_index, kind, place, line, col, label, match, context, hl_start, hl_end, count)
"""
from __future__ import annotations

import os
import re
import time

from .extractors import (KIND_TEXT, ExtractError, extract_units, read_text_file)
from .cache import ExtractCache, pack_units

MAX_HITS_PER_FILE = 5000
CONTEXT_MAX = 300
_DISPLAY_TRANS = str.maketrans({"\n": "↵", "\r": " ", "\t": " ", "\x00": " "})

_rx_cache: dict = {}


def build_regex(opts: dict) -> re.Pattern:
    pat = opts["pattern"]
    if not opts.get("regex"):
        pat = re.escape(pat)
    if opts.get("word"):
        pat = rf"(?<!\w)(?:{pat})(?!\w)"
    flags = re.MULTILINE
    if not opts.get("case"):
        flags |= re.IGNORECASE
    return re.compile(pat, flags)


def _get_regex(opts):
    key = (opts["pattern"], opts.get("regex"), opts.get("word"), opts.get("case"))
    rx = _rx_cache.get(key)
    if rx is None:
        rx = _rx_cache[key] = build_regex(opts)
    return rx


def _open_cache(opts):
    """バッチごとにキャッシュを開く (メイン側が書き込み中でも WAL により並行して読める)"""
    path = opts.get("cache_path") if opts.get("use_cache") else None
    if not path or not os.path.exists(path):
        return None
    try:
        return ExtractCache(path, readonly=True)
    except Exception:
        return None


def _trim(context: str, s: int, e: int) -> tuple[str, int, int]:
    if len(context) <= CONTEXT_MAX:
        return context, s, e
    left = max(0, s - 100)
    right = min(len(context), max(e, left) + CONTEXT_MAX - 100)
    pre = "…" if left > 0 else ""
    post = "…" if right < len(context) else ""
    return pre + context[left:right] + post, s - left + len(pre), min(e, right) - left + len(pre)


def search_units(fidx, kind, units, rx, mode_line: bool, first_only: bool):
    hits = []
    truncated = False
    for unit in units:
        umode, place, base_line, base_col, label, text = unit
        if umode == "L":
            last_line_no = None
            cur_line = base_line
            scan_pos = 0
            for m in rx.finditer(text):
                s, e = m.span()
                if s == e:
                    continue
                cur_line += text.count("\n", scan_pos, s)
                scan_pos = s
                if mode_line and last_line_no == cur_line:
                    h = hits[-1]
                    hits[-1] = h[:10] + (h[10] + 1,)
                    continue
                ls = text.rfind("\n", 0, s) + 1
                le = text.find("\n", s)
                if le < 0:
                    le = len(text)
                ctx = text[ls:le]
                ctx, hs, he = _trim(ctx.translate(_DISPLAY_TRANS), s - ls, min(e, le) - ls)
                hits.append((fidx, kind, place, cur_line, s - ls + 1, label, m.group(0)[:200], ctx, hs, he, 1))
                last_line_no = cur_line
                if first_only or len(hits) >= MAX_HITS_PER_FILE:
                    return hits, len(hits) >= MAX_HITS_PER_FILE
        else:
            first = None
            for m in rx.finditer(text):
                s, e = m.span()
                if s == e:
                    continue
                if mode_line and first is not None:
                    h = hits[-1]
                    hits[-1] = h[:10] + (h[10] + 1,)
                    continue
                if base_line is None:  # 図形など行・列の無い箇所
                    col = None
                else:
                    col = base_col if base_col is not None else s + 1
                ctx, hs, he = _trim(text.translate(_DISPLAY_TRANS), s, e)
                hits.append((fidx, kind, place, base_line, col, label, m.group(0)[:200], ctx, hs, he, 1))
                first = s
                if first_only or len(hits) >= MAX_HITS_PER_FILE:
                    return hits, len(hits) >= MAX_HITS_PER_FILE
    return hits, truncated


def process_batch(batch, opts):
    """batch: [(file_index, path, kind, size, mtime)]
    戻り値: [(file_index, status, message, hits, encoding, packed_units_for_cache, truncated)]
       status: 'ok' | 'cached' | 'error' | 'skip'
    """
    rx = _get_regex(opts)
    mode_line = opts.get("mode", "line") == "line"
    first_only = bool(opts.get("first_only"))
    cache = _open_cache(opts) if any(b[2] != KIND_TEXT for b in batch) else None
    results = []
    for fidx, path, kind, size, mtime in batch:
        enc = ""
        packed = None
        status = "ok"
        try:
            if kind == KIND_TEXT:
                text, enc = read_text_file(path, opts.get("encoding", "auto"))
                units = [("L", "", 1, None, "", text)]
            else:
                units = cache.get(path, mtime, size) if cache else None
                if units is not None:
                    status = "cached"
                else:
                    units = extract_units(path)
                    if opts.get("use_cache"):
                        packed = pack_units(units)
            hits, truncated = search_units(fidx, kind, units, rx, mode_line, first_only)
            results.append((fidx, status, "", hits, enc, packed, truncated))
        except ExtractError as e:
            st = "skip" if "バイナリ" in str(e) else "error"
            results.append((fidx, st, str(e), [], enc, None, False))
        except PermissionError:
            results.append((fidx, "error", "アクセス権が無い、または他のアプリで開かれています", [], enc, None, False))
        except FileNotFoundError:
            results.append((fidx, "error", "ファイルが見つかりません (検索中に削除/移動された可能性)", [], enc, None, False))
        except MemoryError:
            results.append((fidx, "error", "ファイルが大きすぎて読めません", [], enc, None, False))
        except Exception as e:  # 想定外 (壊れたファイル等)
            results.append((fidx, "error", f"読み取りエラー: {type(e).__name__}: {e}", [], enc, None, False))
    if cache is not None:
        cache.close()
    return results


def warmup():
    """プロセスプールの起動確認用 (重いモジュールを先に読み込んでおく)"""
    try:
        import pypdfium2  # noqa: F401
        import xlrd  # noqa: F401
        import olefile  # noqa: F401
    except Exception:
        pass
    return os.getpid(), time.time()
