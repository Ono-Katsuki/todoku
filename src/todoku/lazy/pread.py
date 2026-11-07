"""位置を指定して読む（os.pread）。Windows には os.pread がないので、ファイルの位置を動かして読む（同じ fd を使う読みは錠で順番に）。"""

from __future__ import annotations

import os
import threading

# Windows の os.open は、O_BINARY がないと改行を変える（字の数がずれる）。ほかでは 0
O_BINARY = getattr(os, "O_BINARY", 0)

if hasattr(os, "pread"):
    pread = os.pread
else:
    _lock = threading.Lock()

    def pread(fd: int, n: int, offset: int) -> bytes:
        with _lock:
            os.lseek(fd, offset, os.SEEK_SET)
            return os.read(fd, n)
