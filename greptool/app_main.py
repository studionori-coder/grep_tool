import multiprocessing
import sys


def _selftest(folder: str, pattern: str) -> int:
    """画面を出さずに検索だけ行い、結果をファイルに書く (exe 化後の動作確認用)
    使い方: GrepTool.exe --selftest <フォルダ> <文字列>  → カレントに greptool_selftest.txt
    """
    import threading
    from .engine import SearchEngine

    class CB:
        def enumerating(self, n): pass
        def started(self, files, total): pass
        def progress(self, *a): pass
        def results(self, files, hits): self.hits.extend(hits)

    cb = CB()
    cb.hits = []
    eng = SearchEngine()
    opts = dict(pattern=pattern, regex=False, word=False, case=False, encoding="auto", mode="line",
                first_only=False, use_cache=False, cache_path=None, folders=[folder], include="*.*",
                exclude_files="", exclude_dirs="", recursive=True,
                kinds=["text", "excel", "word", "ppt", "pdf"])
    files, s = eng.run(opts, cb, threading.Event())
    eng.shutdown()
    with open("greptool_selftest.txt", "w", encoding="utf-8") as f:
        f.write(f"files={s.files_total} hits={s.hits} errors={s.errors} elapsed={s.elapsed:.2f}s\n")
        for h in sorted(cb.hits):
            f.write(f"{files[h[0]].path}\t{h[1]}\t{h[2]}\t{h[3]}\t{h[4]}\t{h[5]}\t{h[7]}\n")
        for fi in files:
            if fi.status == "error":
                f.write(f"ERROR\t{fi.path}\t{fi.message}\n")
    return 0


def main():
    multiprocessing.freeze_support()  # exe 化したときの並列検索プロセス用 (必須)
    if len(sys.argv) >= 4 and sys.argv[1] == "--selftest":
        sys.exit(_selftest(sys.argv[2], sys.argv[3]))
    from .gui import run
    sys.exit(run())
