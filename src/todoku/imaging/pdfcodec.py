"""PDF の画像のフィルター JBIG2Decode・JPXDecode を開く入口と、開いた結果の控え。

画像を開くのに数秒〜数十秒かかるので、同じ画像を同じ働き手のプロセスで開き直さない（view_page で範囲を変えて拡大して読むとき）。
控えの鍵はデータのハッシュだけ。暗号化した書類の画像は控えない（解いた中身をプロセスに残さない）。
"""

from __future__ import annotations

import hashlib
from collections import OrderedDict

_CACHE: "OrderedDict[tuple, tuple]" = OrderedDict()
_CACHE_BYTES = 300 * 1024 * 1024


def _size(v) -> int:
    px = getattr(v, "px", None)
    if px is not None:
        return len(px)
    return sum(len(p) for p in v.planes) + (len(v.alpha) if v.alpha else 0)


def _store(key: tuple, value) -> None:
    n = _size(value)
    if n > _CACHE_BYTES // 2:
        return
    _CACHE[key] = (value, n)
    while sum(x[1] for x in _CACHE.values()) > _CACHE_BYTES:
        _CACHE.popitem(last=False)


def jbig2(data: bytes, globals_data: bytes | None, width: int | None, height: int | None, deadline: float | None, cache: bool = True):
    """imaging.jbig2.Page（0＝黒、255＝白）。"""
    from . import jbig2 as J

    h = hashlib.sha256(data)
    h.update(b"\0")
    h.update(globals_data or b"")
    key = ("jbig2", h.hexdigest(), width, height)
    hit = _CACHE.get(key) if cache else None
    if hit:
        _CACHE.move_to_end(key)
        return hit[0]
    page = J.decode(data, globals_data, width, height, deadline=deadline)
    if cache and not page.stopped:
        _store(key, page)
    return page


def jpx(data: bytes, target_pixels: int | None, deadline: float | None, cache: bool = True):
    """imaging.jpx.Image。控えの中に、画面に足りる解像度のものがあれば使う。"""
    from . import jpx as X

    dig = hashlib.sha256(data).hexdigest()
    if cache:
        for key, (img, _n) in reversed(_CACHE.items()):
            if key[:2] == ("jpx", dig) and (img.reduce == 0 or not target_pixels or img.w * img.h >= target_pixels):
                _CACHE.move_to_end(key)
                return img
    img = X.decode(data, target_pixels=target_pixels, deadline=deadline)
    if cache and not img.stopped:
        _store(("jpx", dig, img.reduce), img)
    return img
