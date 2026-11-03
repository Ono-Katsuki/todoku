"""色空間（ISO 32000-1 8.6）と関数（7.10）：色の成分 → 画面の RGB（0〜1）。

- DeviceGray・DeviceRGB・DeviceCMYK（10.3 の単純な変換）、CalGray・CalRGB（Device と同じに扱う）、Lab（D50 → sRGB）、
  ICCBased（成分の数 N の Device の色空間、なければ Alternate）、Indexed（表）、Separation・DeviceN（色の変換の関数で
  代わりの色空間へ。All は黒の濃さ、None は描かない）、Pattern（模様：paint.py が描く）
- 関数：型 0（標本）・型 2（指数）・型 3（つなぎ）・型 4（PostScript の計算）
"""

from __future__ import annotations

import math

from ..objects import Name, Stream

_ABBR = {"G": "DeviceGray", "RGB": "DeviceRGB", "CMYK": "DeviceCMYK", "I": "Indexed"}


class CS:
    """色空間。n：成分の数、rgb(comps) → (r, g, b)（0〜1）。"""

    kind = "DeviceGray"
    n = 1
    none = False  # Separation の None（描かない）

    def rgb(self, c):
        v = _clamp(c[0]) if c else 0.0
        return (v, v, v)

    def initial(self):
        return [0.0] * self.n

    def ranges(self):
        return [(0.0, 1.0)] * self.n


def _clamp(v) -> float:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return 0.0
    return 0.0 if v < 0 else 1.0 if v > 1 else v


class Gray(CS):
    pass


class RGB(CS):
    kind = "DeviceRGB"
    n = 3

    def rgb(self, c):
        c = list(c) + [0.0] * (3 - len(c))
        return (_clamp(c[0]), _clamp(c[1]), _clamp(c[2]))


class CMYK(CS):
    kind = "DeviceCMYK"
    n = 4

    def initial(self):
        return [0.0, 0.0, 0.0, 1.0]

    def rgb(self, c):
        c = list(c) + [0.0] * (4 - len(c))
        cc, m, y, k = (_clamp(v) for v in c[:4])
        return ((1 - cc) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k))


class Lab(CS):
    kind = "Lab"
    n = 3

    def __init__(self, wp=(0.9505, 1.0, 1.089), rng=(-100, 100, -100, 100)):
        self.wp = wp
        self.rng = rng

    def initial(self):
        return [0.0, max(self.rng[0], 0.0), max(self.rng[2], 0.0)]

    def ranges(self):
        return [(0.0, 100.0), (self.rng[0], self.rng[1]), (self.rng[2], self.rng[3])]

    def rgb(self, c):
        c = list(c) + [0.0] * (3 - len(c))
        L, a, b = float(c[0]), float(c[1]), float(c[2])
        fy = (L + 16) / 116
        fx, fz = fy + a / 500, fy - b / 200

        def g(t):
            return t ** 3 if t > 6 / 29 else 3 * (6 / 29) ** 2 * (t - 4 / 29)

        X, Y, Z = self.wp[0] * g(fx), self.wp[1] * g(fy), self.wp[2] * g(fz)
        r = 3.2406 * X - 1.5372 * Y - 0.4986 * Z
        gg = -0.9689 * X + 1.8758 * Y + 0.0415 * Z
        bb = 0.0557 * X - 0.2040 * Y + 1.0570 * Z

        def s(v):
            v = max(0.0, v)
            return _clamp(12.92 * v if v <= 0.0031308 else 1.055 * v ** (1 / 2.4) - 0.055)

        return (s(r), s(gg), s(bb))


class Indexed(CS):
    kind = "Indexed"
    n = 1

    def __init__(self, base: CS, hival: int, lut: bytes):
        self.base, self.hival, self.lut = base, hival, lut

    def ranges(self):
        return [(0.0, float(self.hival))]

    def rgb(self, c):
        i = int(c[0]) if c else 0
        i = max(0, min(self.hival, i))
        nb = self.base.n
        raw = self.lut[i * nb:(i + 1) * nb]
        if len(raw) < nb:
            raw = bytes(raw) + bytes(nb - len(raw))
        rg = self.base.ranges()
        return self.base.rgb([rg[k][0] + raw[k] / 255 * (rg[k][1] - rg[k][0]) for k in range(nb)])


class Tint(CS):
    """Separation・DeviceN：色の変換の関数で代わりの色空間へ。"""

    kind = "Separation"

    def __init__(self, n: int, alt: CS, fn, names=()):
        self.n, self.alt, self.fn = n, alt, fn
        self.none = bool(names) and all(x == "None" for x in names)
        self.all = list(names) == ["All"]
        self._cache: dict = {}

    def initial(self):
        return [1.0] * self.n

    def rgb(self, c):
        key = tuple(round(float(v), 4) for v in c[:self.n])
        hit = self._cache.get(key)
        if hit is not None:
            return hit
        if self.all:
            v = 1 - _clamp(key[0] if key else 1.0)
            out = (v, v, v)
        elif self.fn is not None:
            try:
                out = self.alt.rgb(self.fn(list(key)))
            except Exception:  # noqa: BLE001
                v = 1 - _clamp(sum(key) / max(1, len(key)))
                out = (v, v, v)
        else:
            v = 1 - _clamp(key[0] if key else 1.0)
            out = (v, v, v)
        if len(self._cache) < 4096:
            self._cache[key] = out
        return out


class PatternCS(CS):
    kind = "Pattern"
    n = 0

    def __init__(self, under: CS | None):
        self.under = under
        self.n = under.n if under else 0

    def rgb(self, c):
        return self.under.rgb(c) if self.under else (0.0, 0.0, 0.0)


GRAY, DRGB, DCMYK = Gray(), RGB(), CMYK()


def colorspace(doc, cs, res: dict | None = None, depth: int = 0) -> CS:
    """色空間の名前・配列 → CS。リソースの ColorSpace の名前も引く。分からない色空間は Gray。"""
    cs = doc.resolve(cs)
    if depth > 8:
        return GRAY
    if isinstance(cs, Name):
        n = _ABBR.get(cs.name, cs.name)
        if n == "DeviceGray" or n == "CalGray":
            return GRAY
        if n == "DeviceRGB" or n == "CalRGB":
            return DRGB
        if n == "DeviceCMYK":
            return DCMYK
        if n == "Pattern":
            return PatternCS(None)
        named = doc.resolve((doc.resolve((res or {}).get("ColorSpace")) or {}).get(cs.name)) if res else None
        if named is not None and named is not cs:
            return colorspace(doc, named, None, depth + 1)
        return GRAY
    if isinstance(cs, list) and cs:
        kind = _ABBR.get(getattr(doc.resolve(cs[0]), "name", ""), getattr(doc.resolve(cs[0]), "name", ""))
        if kind in ("CalGray",):
            return GRAY
        if kind in ("CalRGB",):
            return DRGB
        if kind == "Lab" and len(cs) > 1:
            d = doc.resolve(cs[1]) or {}
            wp = [float(doc.resolve(v)) for v in (doc.resolve(d.get("WhitePoint")) or [0.9505, 1.0, 1.089])]
            rg = [float(doc.resolve(v)) for v in (doc.resolve(d.get("Range")) or [-100, 100, -100, 100])]
            return Lab(tuple(wp[:3]), tuple(rg[:4]) if len(rg) >= 4 else (-100, 100, -100, 100))
        if kind == "ICCBased" and len(cs) > 1:
            st = doc.resolve(cs[1])
            d = st.dict if isinstance(st, Stream) else {}
            alt = d.get("Alternate")
            if alt is not None:
                a = colorspace(doc, alt, res, depth + 1)
                if a.n:
                    return a
            n = int(doc.resolve(d.get("N")) or 3)
            return {1: GRAY, 3: DRGB, 4: DCMYK}.get(n, DRGB)
        if kind == "Indexed" and len(cs) >= 4:
            base = colorspace(doc, cs[1], res, depth + 1)
            hival = int(doc.resolve(cs[2]) or 0)
            lut = doc.resolve(cs[3])
            lut = doc.decode(lut) if isinstance(lut, Stream) else bytes(getattr(lut, "raw", lut) or b"")
            return Indexed(base, hival, bytes(lut))
        if kind in ("Separation", "DeviceN") and len(cs) >= 4:
            names = doc.resolve(cs[1])
            names = [getattr(doc.resolve(x), "name", "") for x in (names if isinstance(names, list) else [names])]
            alt = colorspace(doc, cs[2], res, depth + 1)
            try:
                fn = function(doc, cs[3])
            except Exception:  # noqa: BLE001
                fn = None
            return Tint(len(names) if kind == "DeviceN" else 1, alt, fn, names)
        if kind == "Pattern":
            return PatternCS(colorspace(doc, cs[1], res, depth + 1) if len(cs) > 1 else None)
        if kind in ("DeviceGray", "DeviceRGB", "DeviceCMYK"):
            return colorspace(doc, Name(kind), res, depth + 1)
    return GRAY


# ====================================================================== 関数（7.10）
def _nums(doc, v) -> list[float]:
    v = doc.resolve(v)
    return [float(doc.resolve(x)) for x in v] if isinstance(v, list) else []


def function(doc, f, depth: int = 0):
    """関数（辞書・ストリーム・その配列）→ 呼べるもの：入力の並び → 出力の並び。"""
    f = doc.resolve(f)
    if depth > 6:
        raise ValueError("関数が深すぎます")
    if isinstance(f, list):
        fs = [function(doc, x, depth + 1) for x in f]
        return lambda xs: [fn(xs)[0] for fn in fs]
    d = f.dict if isinstance(f, Stream) else f
    if not isinstance(d, dict):
        raise ValueError("関数ではありません")
    t = int(doc.resolve(d.get("FunctionType")) or 0)
    dom = _nums(doc, d.get("Domain")) or [0.0, 1.0]
    rng = _nums(doc, d.get("Range"))

    def clip_in(xs):
        return [min(max(float(x), dom[2 * i]), dom[2 * i + 1]) if 2 * i + 1 < len(dom) else float(x) for i, x in enumerate(xs)]

    def clip_out(ys):
        if not rng:
            return ys
        return [min(max(y, rng[2 * i]), rng[2 * i + 1]) if 2 * i + 1 < len(rng) else y for i, y in enumerate(ys)]

    if t == 2:
        c0 = _nums(doc, d.get("C0")) or [0.0]
        c1 = _nums(doc, d.get("C1")) or [1.0]
        N = float(doc.resolve(d.get("N")) or 1)

        def f2(xs):
            x = clip_in(xs)[0]
            xn = x ** N if x >= 0 or N == int(N) else 0.0
            return clip_out([a + xn * (b - a) for a, b in zip(c0, c1)])
        return f2
    if t == 3:
        fs = [function(doc, x, depth + 1) for x in (doc.resolve(d.get("Functions")) or [])]
        bounds = _nums(doc, d.get("Bounds"))
        enc = _nums(doc, d.get("Encode"))

        def f3(xs):
            x = clip_in(xs)[0]
            k = 0
            while k < len(bounds) and x >= bounds[k]:
                k += 1
            k = min(k, len(fs) - 1)
            lo = dom[0] if k == 0 else bounds[k - 1]
            hi = dom[1] if k >= len(bounds) else bounds[k]
            e0, e1 = (enc[2 * k], enc[2 * k + 1]) if 2 * k + 1 < len(enc) else (0.0, 1.0)
            u = e0 if hi == lo else e0 + (x - lo) * (e1 - e0) / (hi - lo)
            return clip_out(fs[k]([u]))
        return f3
    if t == 0 and isinstance(f, Stream):
        data = doc.decode(f)
        size = [int(v) for v in _nums(doc, d.get("Size"))]
        bps = int(doc.resolve(d.get("BitsPerSample")) or 8)
        m = len(dom) // 2
        n = len(rng) // 2
        enc = _nums(doc, d.get("Encode")) or [v for s in size for v in (0, s - 1)]
        dec = _nums(doc, d.get("Decode")) or rng
        total = 1
        for s in size:
            total *= max(1, s)
        if total * n > 4_000_000:
            raise ValueError("標本の関数が大きすぎます")
        maxv = (1 << bps) - 1

        def sample(idx):
            # idx：各入力の番号 → 出力の並び
            off = 0
            mulk = 1
            for i, v in enumerate(idx):
                off += v * mulk
                mulk *= size[i]
            out = []
            for j in range(n):
                bit = (off * n + j) * bps
                byte = bit >> 3
                if bps == 8:
                    v = data[byte] if byte < len(data) else 0
                elif bps == 16:
                    v = int.from_bytes(data[byte:byte + 2].ljust(2, b"\0"), "big")
                elif bps == 32:
                    v = int.from_bytes(data[byte:byte + 4].ljust(4, b"\0"), "big")
                elif bps == 24:
                    v = int.from_bytes(data[byte:byte + 3].ljust(3, b"\0"), "big")
                else:
                    v = ((data[byte] if byte < len(data) else 0) >> (8 - bps - (bit & 7))) & maxv
                out.append(dec[2 * j] + v * (dec[2 * j + 1] - dec[2 * j]) / maxv)
            return out

        def f0(xs):
            x = clip_in(xs)
            # 多次元は最も近い標本、1 次元は直線で補う
            e = []
            for i in range(m):
                lo, hi = dom[2 * i], dom[2 * i + 1]
                u = enc[2 * i] + (x[i] - lo) * (enc[2 * i + 1] - enc[2 * i]) / ((hi - lo) or 1)
                e.append(min(max(u, 0), size[i] - 1))
            if m == 1:
                i0 = int(math.floor(e[0]))
                i1 = min(i0 + 1, size[0] - 1)
                t_ = e[0] - i0
                a, b = sample([i0]), sample([i1])
                return clip_out([p + (q - p) * t_ for p, q in zip(a, b)])
            return clip_out(sample([int(round(v)) for v in e]))
        return f0
    if t == 4 and isinstance(f, Stream):
        prog = _ps_parse(doc.decode(f))

        def f4(xs):
            st = list(clip_in(xs))
            _ps_run(prog, st, 0)
            return clip_out([float(v) for v in st[-(len(rng) // 2):]] if rng else [float(v) for v in st])
        return f4
    raise ValueError(f"関数の型 {t} は扱えません")


def _ps_parse(data: bytes):
    toks = data.replace(b"{", b" { ").replace(b"}", b" } ").split()
    pos = 0

    def block():
        nonlocal pos
        out = []
        while pos < len(toks):
            t = toks[pos]
            pos += 1
            if t == b"{":
                out.append(block())
            elif t == b"}":
                return out
            else:
                try:
                    out.append(float(t) if b"." in t or b"e" in t.lower() else int(t))
                except ValueError:
                    out.append(t.decode("latin-1"))
        return out

    while pos < len(toks) and toks[pos] != b"{":
        pos += 1
    pos += 1
    return block()


def _ps_run(prog, st, depth):
    if depth > 20:
        raise ValueError("深すぎます")
    i = 0
    n = len(prog)
    steps = 0
    while i < n:
        steps += 1
        if steps > 100000:
            raise ValueError("長すぎます")
        t = prog[i]
        i += 1
        if isinstance(t, (int, float)):
            st.append(t)
            continue
        if isinstance(t, list):
            # if / ifelse の手続き
            if i < n and prog[i] == "if":
                i += 1
                if st.pop():
                    _ps_run(t, st, depth + 1)
                continue
            if i + 1 < n and isinstance(prog[i], list) and prog[i + 1] == "ifelse":
                other = prog[i]
                i += 2
                _ps_run(t if st.pop() else other, st, depth + 1)
                continue
            continue
        op = t
        if op == "true":
            st.append(True)
        elif op == "false":
            st.append(False)
        elif op == "dup":
            st.append(st[-1])
        elif op == "pop":
            st.pop()
        elif op == "exch":
            st[-1], st[-2] = st[-2], st[-1]
        elif op == "copy":
            k = int(st.pop())
            if k:
                st.extend(st[-k:])
        elif op == "index":
            k = int(st.pop())
            st.append(st[-1 - k])
        elif op == "roll":
            j, k = int(st.pop()), int(st.pop())
            if k:
                seg = st[-k:]
                j %= k
                st[-k:] = seg[-j:] + seg[:-j] if j else seg
        elif op in _PS2:
            b, a = st.pop(), st.pop()
            st.append(_PS2[op](a, b))
        elif op in _PS1:
            st.append(_PS1[op](st.pop()))
        else:
            raise ValueError(f"PostScript の {op} は扱えません")


def _idiv(a, b):
    return int(a / b) if b else 0


_PS2 = {"add": lambda a, b: a + b, "sub": lambda a, b: a - b, "mul": lambda a, b: a * b,
        "div": lambda a, b: a / b if b else 0.0, "idiv": _idiv, "mod": lambda a, b: int(math.fmod(a, b)) if b else 0,
        "atan": lambda a, b: math.degrees(math.atan2(a, b)) % 360, "exp": lambda a, b: a ** b if a or b > 0 else 0.0,
        "eq": lambda a, b: a == b, "ne": lambda a, b: a != b, "gt": lambda a, b: a > b, "ge": lambda a, b: a >= b,
        "lt": lambda a, b: a < b, "le": lambda a, b: a <= b,
        "and": lambda a, b: (a and b) if isinstance(a, bool) else int(a) & int(b),
        "or": lambda a, b: (a or b) if isinstance(a, bool) else int(a) | int(b),
        "xor": lambda a, b: (a != b) if isinstance(a, bool) else int(a) ^ int(b),
        "bitshift": lambda a, b: int(a) << int(b) if b >= 0 else int(a) >> -int(b)}
_PS1 = {"abs": abs, "neg": lambda a: -a, "ceiling": math.ceil, "floor": math.floor, "round": lambda a: math.floor(a + 0.5),
        "truncate": math.trunc, "sqrt": lambda a: math.sqrt(max(0, a)), "sin": lambda a: math.sin(math.radians(a)),
        "cos": lambda a: math.cos(math.radians(a)), "ln": lambda a: math.log(a) if a > 0 else 0.0,
        "log": lambda a: math.log10(a) if a > 0 else 0.0, "cvi": lambda a: int(a), "cvr": float,
        "not": lambda a: (not a) if isinstance(a, bool) else ~int(a)}

