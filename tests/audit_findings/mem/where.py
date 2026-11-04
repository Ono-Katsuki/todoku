"""監査用：どこで増えているか。1 つの書類を 1 つの道具で動かし、常駐が X MB を超えるか T 秒たった所で、スタックを出して終わる。

使い方：where.py 書類 道具 [MB=1000] [秒=20]
親が 1.5 GB で殺す見張りを続ける（子のスレッドが GIL を取れないときのため）。
"""
import fcntl
import faulthandler
import os
import resource
import subprocess
import sys
import tempfile
import threading
import time

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.abspath(os.path.join(HERE, "..", "..", "..", "src"))
LOCK = os.path.join(tempfile.gettempdir(), "todoku_mem_guard.lock")


def child(path, tool, mb, secs):
    os.environ["TODOKU_ISOLATE"] = "0"
    sys.argv = ["probe.py", path, tool, tempfile.gettempdir()]

    def watch():
        t0 = time.time()
        while True:
            time.sleep(0.05)
            rss = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1048576
            if rss > mb or time.time() - t0 > secs:
                print(f"--- rss={rss:.0f}MB t={time.time() - t0:.1f}s", flush=True)
                faulthandler.dump_traceback(file=sys.stdout, all_threads=True)
                sys.stdout.flush()
                os._exit(9)
    threading.Thread(target=watch, daemon=True).start()
    sys.path.insert(0, HERE)
    import probe

    probe.main()


def rss_kb(pid):
    try:
        return int(subprocess.run(["ps", "-o", "rss=", "-p", str(pid)], capture_output=True, text=True).stdout.strip() or 0)
    except ValueError:
        return 0


def main():
    if sys.argv[1] == "--child":
        return child(sys.argv[2], sys.argv[3], float(sys.argv[4]), float(sys.argv[5]))
    path, tool = sys.argv[1], sys.argv[2]
    mb = sys.argv[3] if len(sys.argv) > 3 else "1000"
    secs = sys.argv[4] if len(sys.argv) > 4 else "20"
    env = dict(os.environ, PYTHONPATH=SRC, TODOKU_ISOLATE="0", PYTHONUNBUFFERED="1")
    with open(LOCK, "w") as lk:
        fcntl.flock(lk, fcntl.LOCK_EX)
        p = subprocess.Popen(["nice", "-n", "10", sys.executable, os.path.abspath(__file__), "--child", path, tool, mb, secs], env=env,
                             stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, errors="replace")
        t0 = time.time()
        while p.poll() is None:
            if rss_kb(p.pid) > 1.5 * 1048576 or time.time() - t0 > 60:
                p.kill()
                print("(親が殺した：スタックは取れなかった)")
                break
            time.sleep(0.03)
        out = p.communicate()[0]
    import re

    for ln in out.splitlines():
        if ln.startswith("---") or ln.startswith("{"):
            print(ln[:120])
    keep = [re.sub(r'File ".*?/src/todoku/(.*?)", line (\d+) in (.*)', r"\1:\2 \3", ln.strip()) for ln in out.splitlines() if "/src/todoku/" in ln]
    print("\n".join(keep[:int(os.environ.get("WHERE_LINES", 8))]))


if __name__ == "__main__":
    main()
