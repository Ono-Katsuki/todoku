"""JPEG（ITU-T T.81）のプログレッシブ（SOF2）・成分を分けた走査の順次（SOF0・SOF1）・12 ビットの精度・算術符号（SOF9・SOF10、jpegarith.py）・
可逆（SOF3、jpegarith.py）。

走査ごとに係数を積み上げてから、逆 DCT する（pdf/raster.py の部品を使う）。算術符号の可逆（SOF11）と階層（SOF5〜7・13〜15）は扱えない。
係数は成分ごとの平らな array('h')。大きな画像は raster と同じく、逆 DCT を 1/2・1/4・1/8 にする。1/8 なら AC の走査を読まない。
"""

from __future__ import annotations

import struct
from array import array

from ..limits import MAX_PIXELS, LimitExceeded


def decode(data: bytes, max_pixels: int, want_rgb: bool, window: tuple | None = None):
    from ..pdf.raster import RasterError, _Huff, _idct_matrix, _jpeg_rgb, _put_block, _segments, _ZZ, jpeg_blocks, jpeg_window

    pos = 2
    qt: dict = {}
    hts: dict = {}
    frame = None
    restart = 0
    adobe = None
    progressive = False
    arith = lossless = False
    precision = 8
    dac: dict = {}
    coefs: list = []
    n = len(data)
    n_scale = 8
    hmax = vmax = 1
    while pos + 2 <= n:
        if data[pos] != 0xFF:
            pos += 1
            continue
        m = data[pos + 1]
        pos += 2
        if m in (0xD8, 0x01, 0xFF, 0x00) or 0xD0 <= m <= 0xD7:
            if m == 0xFF:
                pos -= 1
            continue
        if m == 0xD9:
            break
        if pos + 2 > n:
            break
        ln = struct.unpack_from(">H", data, pos)[0]
        seg = data[pos + 2:pos + ln]
        if m == 0xDB:
            i = 0
            while i < len(seg):
                pq, tq = seg[i] >> 4, seg[i] & 15
                i += 1
                if pq:
                    qt[tq] = list(struct.unpack_from(">64H", seg, i))
                    i += 128
                else:
                    qt[tq] = list(seg[i:i + 64])
                    i += 64
        elif m == 0xC4:
            i = 0
            while i < len(seg):
                tc, th = seg[i] >> 4, seg[i] & 15
                counts = seg[i + 1:i + 17]
                k = sum(counts)
                hts[(tc, th)] = _Huff(counts, seg[i + 17:i + 17 + k])
                i += 17 + k
        elif m in (0xC0, 0xC1, 0xC2, 0xC3, 0xC9, 0xCA):
            progressive = m in (0xC2, 0xCA)
            arith = m in (0xC9, 0xCA)
            lossless = m == 0xC3
            p, hh, ww, nf = seg[0], *struct.unpack_from(">HH", seg, 1), seg[5]
            if (lossless and not 2 <= p <= 16) or (not lossless and p not in (8, 12)):
                raise RasterError(f"JPEG：{p} ビットの精度は規格にありません")
            precision = p
            comps = []
            for j in range(nf):
                cid, hv, tq = seg[6 + 3 * j], seg[7 + 3 * j], seg[8 + 3 * j]
                comps.append({"id": cid, "h": hv >> 4, "v": hv & 15, "tq": tq})
            if not comps or not all(c["h"] and c["v"] for c in comps):
                raise RasterError("JPEG：成分の間引きの値が 0 です")
            if ww * hh > MAX_PIXELS:
                raise LimitExceeded(f"JPEG の画像（{ww:,} × {hh:,} 画素）が、開ける大きさ {MAX_PIXELS:,} 画素を超えます")
            hmax = max(c["h"] for c in comps)
            vmax = max(c["v"] for c in comps)
            mcux = (ww + 8 * hmax - 1) // (8 * hmax)
            mcuy = (hh + 8 * vmax - 1) // (8 * vmax)
            for c in comps:
                c["bw"], c["bh"] = mcux * c["h"], mcuy * c["v"]
                c["nx"] = ((ww * c["h"] + hmax - 1) // hmax + 7) // 8
                c["ny"] = ((hh * c["v"] + vmax - 1) // vmax + 7) // 8
                if lossless:
                    c["sw"], c["sh"] = (ww * c["h"] + hmax - 1) // hmax, (hh * c["v"] + vmax - 1) // vmax
                    coefs.append((array("H", bytes(c["sw"] * c["sh"] * 2)), c["sw"], c["sh"]))
                else:
                    coefs.append(array("h", bytes(c["bw"] * c["bh"] * 128)) if p == 8 else array("i", bytes(c["bw"] * c["bh"] * 256)))
            frame = (ww, hh, comps)
            while n_scale > 1 and (ww * n_scale // 8) * (hh * n_scale // 8) > max_pixels:
                n_scale //= 2
        elif m == 0xCC:  # DAC：算術符号の条件づけ（B.2.4.3）
            for i in range(0, len(seg) - 1, 2):
                tc, tb, cs = seg[i] >> 4, seg[i] & 15, seg[i + 1]
                dac[(tc, tb)] = cs if tc else (cs & 15, cs >> 4)
        elif m == 0xDD:
            restart = struct.unpack_from(">H", seg, 0)[0]
        elif m == 0xEE and seg[:5] == b"Adobe":
            adobe = seg[11] if len(seg) > 11 else 0
        elif m == 0xDA:
            if frame is None:
                raise RasterError("JPEG：フレームの前に走査があります")
            ns = seg[0]
            sel = []
            for j in range(ns):
                cid, tt = seg[1 + 2 * j], seg[2 + 2 * j]
                ci = next((i for i, c in enumerate(frame[2]) if c["id"] == cid), None)
                if ci is None:
                    raise RasterError("JPEG：走査の成分がフレームにありません")
                sel.append((ci, tt >> 4, tt & 15))
            ss, se, ah_al = seg[1 + 2 * ns], seg[2 + 2 * ns], seg[3 + 2 * ns]
            ah, al = ah_al >> 4, ah_al & 15
            segs = _segments(data, pos + ln)
            consumed = _next_marker(data, pos + ln)
            if lossless:
                from .jpegarith import lossless_scan

                lossless_scan(segs, frame, coefs, [(s_[0], s_[1]) for s_ in sel], hts, restart, ss, al, precision, hmax, vmax, RasterError)
            elif arith:
                from .jpegarith import scan as arith_scan

                if not (n_scale == 1 and progressive and ss > 0):
                    arith_scan(segs, frame, coefs, sel, restart, progressive, ss, se if progressive else 63, ah, al, hmax, vmax, dac)
            elif not (n_scale == 1 and progressive and ss > 0):
                _scan(segs, frame, coefs, sel, hts, restart, progressive, ss, se, ah, al, hmax, vmax, _ZZ, RasterError)
            pos = consumed
            continue
        elif 0xC3 <= m <= 0xCF and m not in (0xC4, 0xC8, 0xCC):
            raise RasterError(f"JPEG：この符号化（SOF{m - 0xC0}）は開けません（ベースライン・拡張の順次・プログレッシブ・可逆のハフマン・算術符号だけ）")
        pos += ln
    if frame is None:
        raise RasterError("JPEG：画像のデータがありません")
    ww, hh, comps = frame
    if lossless:
        return _lossless_out(frame, coefs, precision, al, adobe, hmax, vmax, want_rgb, window)
    center = 1 << (precision - 1)
    shift = precision - 8
    mat = _idct_matrix(n_scale)
    planes = []
    win = jpeg_window(window, ww, hh, hmax, vmax) if window is not None and want_rgb else None
    for ci, c in enumerate(comps):
        bw, bh = c["bw"], c["bh"]
        plane = bytearray(b"\x80" * (bw * n_scale) * (bh * n_scale))
        qv = qt.get(c["tq"])
        if qv is None:
            raise RasterError("JPEG：量子化の表がありません")
        cf = coefs[ci]
        zz = _ZZ
        stride = bw * n_scale
        ax0, ay0, ax1, ay1 = jpeg_blocks(win, c, hmax, vmax) if win else (0, 0, bw, bh)  # 範囲に要るブロックだけ逆 DCT する
        for by in range(ay0, min(bh, ay1)):
            base_row = by * bw * 64
            for bx in range(ax0, min(bw, ax1)):
                base = base_row + bx * 64
                block = cf[base:base + 64]
                c8 = [0] * 64
                if n_scale == 1:
                    c8[0] = block[0] * qv[0]
                else:
                    for k, v in enumerate(block):
                        if v:
                            c8[zz[k]] = v * qv[k]
                if shift:
                    _put_block_hi(c8, mat, n_scale, plane, stride, bx * n_scale, by * n_scale, center, shift)
                else:
                    _put_block(c8, mat, n_scale, plane, stride, bx * n_scale, by * n_scale)
        planes.append((c, plane, stride))
    W, H = ww * n_scale // 8 or 1, hh * n_scale // 8 or 1
    from ..pdf.raster import Gray

    rgb = _jpeg_rgb(planes, hmax, vmax, W, H, adobe, win, n_scale)
    if want_rgb:
        return rgb
    g = rgb.px[0::3] if len(comps) == 1 else bytes((r * 299 + gg * 587 + b * 114) // 1000 for r, gg, b in zip(rgb.px[0::3], rgb.px[1::3], rgb.px[2::3]))
    return Gray(W, H, g)


def _put_block_hi(coef, mat, n, plane, stride, ox, oy, center: int, shift: int) -> None:
    """12 ビットの逆 DCT：値を 8 ビットに丸めて面へ（画面に見せるため）。"""
    top = 2 * center - 1
    half = 1 << (shift - 1)
    tmp = [[sum(mat[x][u] * coef[v * 8 + u] for u in range(n)) for x in range(n)] for v in range(n)]
    for y in range(n):
        o = (oy + y) * stride + ox
        row = bytearray(n)
        for x in range(n):
            s = int(sum(mat[y][v] * tmp[v][x] for v in range(n)) + center + 0.5)
            s = 0 if s < 0 else top if s > top else s
            b = (s + half) >> shift
            row[x] = 255 if b > 255 else b
        plane[o:o + n] = row


def _lossless_out(frame, planes_in, precision: int, pt: int, adobe, hmax: int, vmax: int, want_rgb: bool, window):
    """可逆の面 → 8 ビットの面 → RGB。8 ビットより細かい値は上の桁、少ないビットは広げる（画面に見せるため）。"""
    from ..pdf.raster import Gray, _jpeg_rgb, jpeg_window

    ww, hh, comps = frame
    planes = []
    for c, (arr, w, _h) in zip(comps, planes_in):
        if precision - pt == 8 and pt == 0:
            bytes8 = bytes(arr.tolist())
        elif precision >= 8:
            sh = precision - 8
            bytes8 = bytes([min(255, (v << pt) >> sh) for v in arr])
        else:
            top = (1 << precision) - 1
            bytes8 = bytes([min(255, ((v << pt) * 255 + top // 2) // top) for v in arr])
        planes.append((c, bytearray(bytes8), w))
    win = jpeg_window(window, ww, hh, hmax, vmax) if window is not None and want_rgb else None
    rgb = _jpeg_rgb(planes, hmax, vmax, ww, hh, adobe, win, 8)
    if want_rgb:
        return rgb
    g = rgb.px[0::3] if len(comps) == 1 else bytes((r * 299 + gg * 587 + b * 114) // 1000 for r, gg, b in zip(rgb.px[0::3], rgb.px[1::3], rgb.px[2::3]))
    return Gray(rgb.w, rgb.h, g)


def _next_marker(data: bytes, pos: int) -> int:
    """走査のデータの次のマーカーの位置（RST と詰め物は飛ばす）。"""
    n = len(data)
    while True:
        j = data.find(b"\xff", pos)
        if j < 0 or j + 1 >= n:
            return n
        m = data[j + 1]
        if m == 0x00 or 0xD0 <= m <= 0xD7:
            pos = j + 2
            continue
        if m == 0xFF:
            pos = j + 1
            continue
        return j


def _scan(segs, frame, coefs, sel, hts, restart, progressive, ss, se, ah, al, hmax, vmax, ZZ, RasterError):
    ww, hh, comps = frame
    # 1 つの成分だけの走査は、ブロックが単位（A.2.2）。いくつかなら MCU
    if len(sel) == 1:
        ci = sel[0][0]
        c = comps[ci]
        order = None
        total = c["nx"] * c["ny"]
    else:
        mcux = (ww + 8 * hmax - 1) // (8 * hmax)
        mcuy = (hh + 8 * vmax - 1) // (8 * vmax)
        total = mcux * mcuy
        order = mcux
    per = restart or total
    mi = 0
    p1 = 1 << al
    m1 = -(1 << al)
    for buf in segs:
        if mi >= total:
            break
        buf = buf + b"\0" * 8
        pos = 0
        pred = [0] * len(comps)
        eobrun = 0
        for _ in range(min(per, total - mi)):
            if order is None:
                ci, td, ta = sel[0]
                by, bx = divmod(mi, comps[ci]["nx"])
                blocks = ((ci, td, ta, by * comps[ci]["bw"] * 64 + bx * 64),)
            else:
                my, mx = divmod(mi, order)
                blocks = []
                for ci, td, ta in sel:
                    c = comps[ci]
                    for v in range(c["v"]):
                        for h in range(c["h"]):
                            blocks.append((ci, td, ta, ((my * c["v"] + v) * c["bw"] + mx * c["h"] + h) * 64))
            mi += 1
            for ci, td, ta, base in blocks:
                cf = coefs[ci]
                if not progressive or ss == 0:
                    # DC
                    if ah == 0:
                        dcl = hts.get((0, td))
                        if dcl is None:
                            raise RasterError("JPEG：DC のハフマン表がありません")
                        i = pos >> 3
                        e = dcl.lut[(int.from_bytes(buf[i:i + 4], "big") << (pos & 7) & 0xFFFFFFFF) >> 16]
                        if not e:
                            raise RasterError("JPEG：ハフマン符号が読めません")
                        t = e & 255
                        pos += e >> 8
                        if t:
                            i = pos >> 3
                            v = (int.from_bytes(buf[i:i + 4], "big") >> (32 - (pos & 7) - t)) & ((1 << t) - 1)
                            pos += t
                            pred[ci] += v - (1 << t) + 1 if v < (1 << (t - 1)) else v
                        cf[base] = pred[ci] << al if progressive else pred[ci]
                    else:
                        if (buf[pos >> 3] >> (7 - (pos & 7))) & 1:
                            cf[base] |= p1
                        pos += 1
                    if progressive:
                        continue
                acl = hts.get((1, ta))
                if acl is None:
                    raise RasterError("JPEG：AC のハフマン表がありません")
                lut = acl.lut
                if not progressive:
                    k = 1
                    while k <= 63:
                        i = pos >> 3
                        e = lut[(int.from_bytes(buf[i:i + 4], "big") << (pos & 7) & 0xFFFFFFFF) >> 16]
                        if not e:
                            raise RasterError("JPEG：ハフマン符号が読めません")
                        pos += e >> 8
                        sz = e & 15
                        r = (e & 255) >> 4
                        if not sz:
                            if r != 15:
                                break
                            k += 16
                            continue
                        k += r
                        if k > 63:
                            break
                        i = pos >> 3
                        v = (int.from_bytes(buf[i:i + 4], "big") >> (32 - (pos & 7) - sz)) & ((1 << sz) - 1)
                        pos += sz
                        cf[base + k] = v - (1 << sz) + 1 if v < (1 << (sz - 1)) else v
                        k += 1
                    continue
                if ah == 0:
                    # AC の最初の走査（G.1.2.2）
                    if eobrun:
                        eobrun -= 1
                        continue
                    k = ss
                    while k <= se:
                        i = pos >> 3
                        e = lut[(int.from_bytes(buf[i:i + 4], "big") << (pos & 7) & 0xFFFFFFFF) >> 16]
                        if not e:
                            raise RasterError("JPEG：ハフマン符号が読めません")
                        pos += e >> 8
                        sz = e & 15
                        r = (e & 255) >> 4
                        if sz:
                            k += r
                            if k > 63:
                                break
                            i = pos >> 3
                            v = (int.from_bytes(buf[i:i + 4], "big") >> (32 - (pos & 7) - sz)) & ((1 << sz) - 1)
                            pos += sz
                            cf[base + k] = (v - (1 << sz) + 1 if v < (1 << (sz - 1)) else v) * (1 << al)
                            k += 1
                        else:
                            if r < 15:
                                eobrun = (1 << r) - 1
                                if r:
                                    i = pos >> 3
                                    eobrun += (int.from_bytes(buf[i:i + 4], "big") >> (32 - (pos & 7) - r)) & ((1 << r) - 1)
                                    pos += r
                                break
                            k += 16
                    continue
                # AC を細かくする走査（G.1.2.3）
                k = ss
                if eobrun == 0:
                    while k <= se:
                        i = pos >> 3
                        e = lut[(int.from_bytes(buf[i:i + 4], "big") << (pos & 7) & 0xFFFFFFFF) >> 16]
                        if not e:
                            raise RasterError("JPEG：ハフマン符号が読めません")
                        pos += e >> 8
                        sz = e & 15
                        r = (e & 255) >> 4
                        newv = 0
                        if sz:
                            newv = p1 if (buf[pos >> 3] >> (7 - (pos & 7))) & 1 else m1
                            pos += 1
                        elif r < 15:
                            eobrun = 1 << r
                            if r:
                                i = pos >> 3
                                eobrun += (int.from_bytes(buf[i:i + 4], "big") >> (32 - (pos & 7) - r)) & ((1 << r) - 1)
                                pos += r
                            break
                        while k <= se:
                            q = base + k
                            cv = cf[q]
                            if cv:
                                if (buf[pos >> 3] >> (7 - (pos & 7))) & 1 and not cv & p1:
                                    cf[q] = cv + (p1 if cv >= 0 else m1)
                                pos += 1
                            else:
                                if r == 0:
                                    if newv:
                                        cf[q] = newv
                                    k += 1
                                    break
                                r -= 1
                            k += 1
                if eobrun > 0:
                    while k <= se:
                        q = base + k
                        cv = cf[q]
                        if cv:
                            if (buf[pos >> 3] >> (7 - (pos & 7))) & 1 and not cv & p1:
                                cf[q] = cv + (p1 if cv >= 0 else m1)
                            pos += 1
                        k += 1
                    eobrun -= 1
