"""監査用：tests/fixtures/public の小さな書類の「宣言された数」を巨大な値に変えた変異体と、ランダムなビット反転を作る。

使い方：mutants.py 出力フォルダ   （出力フォルダ/MANIFEST.jsonl に 1 行ずつ）
すべて種（seed）固定で同じものができる。
"""
import io
import json
import os
import random
import re
import struct
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
sys.path.insert(0, os.path.join(ROOT, "src"))
PUB = os.path.join(ROOT, "tests", "fixtures", "public")
HUGE = [b"1048576", b"4294967295"]
ATTR = re.compile(rb'([\w:.-]+)="(\d{1,12})"')
REF = re.compile(rb'\b(ref|sqref|spans)="([A-Z0-9:$ ]+)"')


def rezip(blob: bytes, changes: dict) -> bytes:
    src = zipfile.ZipFile(io.BytesIO(blob))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for i in src.infolist():
            data = changes.get(i.filename, src.read(i.filename))
            ni = zipfile.ZipInfo(i.filename, i.date_time)
            ni.compress_type = i.compress_type
            z.writestr(ni, data)
    return out.getvalue()


def zip_mutants(name: str, blob: bytes, rng):
    z = zipfile.ZipFile(io.BytesIO(blob))
    for part in z.namelist():
        if not part.endswith((".xml", ".rels", ".hpf")):
            continue
        d = z.read(part)
        seen = set()
        for m in ATTR.finditer(d):
            a = m[1]
            if a in seen or a.startswith(b"xmlns"):
                continue
            seen.add(a)
            for hv in HUGE:
                yield f"attr:{part}:{a.decode()}={hv.decode()}:all", rezip(blob, {part: re.sub(rb'(\b' + re.escape(a) + rb'=")\d{1,12}(")', rb"\g<1>" + hv + rb"\2", d)})
            nd = d[:m.start(2)] + b"2147483647" + d[m.end(2):]
            yield f"attr:{part}:{a.decode()}=2147483647:first", rezip(blob, {part: nd})
        for m in REF.finditer(d):
            k = (m[1], m[2])
            if k in seen:
                continue
            seen.add(k)
            hv = b"1:1048576" if m[1] == b"spans" else b"A1:XFD1048576"
            yield f"ref:{part}:{m[1].decode()}={hv.decode()}", rezip(blob, {part: d[:m.start(2)] + hv + d[m.end(2):]})
        # 部品の中のビット反転（ZIP の CRC は作り直す）
        for k in range(4):
            b = bytearray(d)
            for _ in range(rng.randint(1, 6)):
                if b:
                    b[rng.randrange(len(b))] ^= 1 << rng.randrange(8)
            yield f"partflip:{part}:{k}", rezip(blob, {part: bytes(b)})


def flips(name: str, blob: bytes, rng, n=15):
    for k in range(n):
        b = bytearray(blob)
        for _ in range(rng.choice([1, 1, 2, 4, 16])):
            b[rng.randrange(len(b))] ^= 1 << rng.randrange(8)
        yield f"flip:{k}", bytes(b)


# ---------------------------------------------------------------------- PDF
OBJ = re.compile(rb"(\d+) (\d+) obj\b(.*?)\bendobj", re.S)


def reflow_pdf(blob: bytes, edit=None) -> bytes:
    """目次（xref）を作り直す：数の長さを変えても位置がずれないように。"""
    tr = re.search(rb"trailer\s*(<<.*?>>)\s*startxref", blob, re.S)
    head = blob[:blob.index(b"\n", 9) + 1]
    out = bytearray(head)
    offs = {}
    for m in OBJ.finditer(blob):
        n = int(m[1])
        body = m[3]
        offs[n] = len(out)
        out += b"%d %s obj" % (n, m[2]) + body + b"endobj\n"
    x = len(out)
    size = max(offs) + 1
    out += b"xref\n0 %d\n0000000000 65535 f \n" % size
    for n in range(1, size):
        out += (b"%010d 00000 n \n" % offs[n]) if n in offs else b"0000000000 00000 f \n"
    out += b"trailer\n" + (tr[1] if tr else b"<< /Size %d /Root 1 0 R >>" % size) + b"\nstartxref\n%d\n%%%%EOF\n" % x
    return bytes(out)


KEYNUM = re.compile(rb"/([A-Za-z]+)\s+(\d+(?:\.\d+)?)(?![\d ]* (?:R|obj))\b")
ARRAY = re.compile(rb"/([A-Za-z]+)\s*\[\s*([\d. ]+?)\s*\]")


def pdf_mutants(name: str, blob: bytes, rng):
    seen = set()
    for m in OBJ.finditer(blob):
        body = m[3]
        cut = body.find(b"stream")
        dic = body if cut < 0 else body[:cut]
        for km in KEYNUM.finditer(dic):
            if km[1] in seen or km[1] == b"Length":
                continue
            seen.add(km[1])
            for hv in (b"1048576", b"4294967295", b"1e9"):
                nb = blob[:m.start(3) + km.start(2)] + hv + blob[m.start(3) + km.end(2):]
                yield f"pdfkey:{km[1].decode()}={hv.decode()}", reflow_pdf(nb)
        for am in ARRAY.finditer(dic):
            if am[1] in seen:
                continue
            seen.add(am[1])
            nums = am[2].split()
            for i in range(min(len(nums), 6)):
                nn = list(nums)
                nn[i] = b"4294967295" if i % 2 else b"1000000000"
                nb = blob[:m.start(3) + am.start(2)] + b" ".join(nn) + blob[m.start(3) + am.end(2):]
                yield f"pdfarr:{am[1].decode()}[{i}]huge", reflow_pdf(nb)
            nb = blob[:m.start(3) + am.start(2)] + b" ".join([b"1000000000"] * len(nums)) + blob[m.start(3) + am.end(2):]
            yield f"pdfarr:{am[1].decode()}:allhuge", reflow_pdf(nb)
        # 中身のストリーム（圧縮されていないとき）の数をところどころ巨大に
        if cut >= 0 and b"/Filter" not in dic and b"/Length" in dic:
            s = body[cut:]
            if re.search(rb"\b(Tf|Td|re|cm|Tm|TJ|Tj)\b", s):
                for k in range(3):
                    def bump(mm):
                        return b"1000000000" if rng.random() < 0.15 else mm[0]
                    ns = re.sub(rb"(?<![\w.])\d+(?:\.\d+)?(?![\w.])", bump, s)
                    # /Length は直さない：読み手の修復にも当てる
                    nb = blob[:m.start(3) + cut] + ns + blob[m.start(3) + len(body):]
                    yield f"pdfcontent:{int(m[1])}:{k}", reflow_pdf(nb)


# ---------------------------------------------------------------------- 複合ファイル（.xls・.doc・.jtd）
def cfb_mutants(name: str, blob: bytes, rng):
    from todoku.cfb import CFB

    c = CFB(blob)
    # 見出し
    for off in [0x1E, 0x20, 0x2C, 0x30, 0x34, 0x38, 0x3C, 0x40, 0x44, 0x48, 0x4C, 0x50]:
        for hv in (0x7FFFFFFF, 0xFFFFFFF0, 0x00100000):
            b = bytearray(blob)
            if off in (0x1E, 0x20):
                struct.pack_into("<H", b, off, min(hv, 0xFFFF) if hv < 0x100000 else 30)
            else:
                struct.pack_into("<I", b, off, hv)
            yield f"cfbhead:{off:#x}={hv:#x}", bytes(b)
    # 目録の大きさ・始まり（目録のセクタが連続しているものとして、生のバイトで探す）
    dir_off = c._off(c.dir_start)
    for k in range(min(6, c.ssize // 128)):
        for fo, hv in ((0x78, 0x7FFFFFFF), (0x78, 0xFFFFFFF0), (0x74, 0x7FFFFFF0), (0x44, 0xFFFFFFF0), (0x4C, 5)):
            b = bytearray(blob)
            struct.pack_into("<I", b, dir_off + k * 128 + fo, hv)
            yield f"cfbdir:{k}:{fo:#x}={hv:#x}", bytes(b)
    # 流れ（ストリーム）の中身：生のバイトの中で連続している所だけ（小さな書類なので、たいてい連続）
    for path in list(c.streams()):
        try:
            data = c.read(path)
        except Exception:  # noqa: BLE001
            continue
        at = blob.find(data[:256]) if len(data) >= 16 else -1
        if at < 0 or blob[at:at + len(data)] != data:
            continue
        if path.endswith("Workbook") or path.endswith("Book"):
            yield from biff_mutants(blob, data, at, rng)
        else:
            words = list(range(0, len(data) - 3, 4))
            rng.shuffle(words)
            for off in words[:50]:
                for hv in (0xFFFFFFFF, 0x7FFFFFFF):
                    b = bytearray(blob)
                    struct.pack_into("<I", b, at + off, hv)
                    yield f"stream:{path}:{off:#x}={hv:#x}", bytes(b)
            hw = list(range(0, len(data) - 1, 2))
            rng.shuffle(hw)
            for off in hw[:25]:
                b = bytearray(blob)
                struct.pack_into("<H", b, at + off, 0xFFFF)
                yield f"stream16:{path}:{off:#x}=ffff", bytes(b)


def biff_mutants(blob, data, at, rng):
    pos = 0
    seen = set()
    while pos + 4 <= len(data):
        op, ln = struct.unpack_from("<HH", data, pos)
        if op not in seen:
            seen.add(op)
            for f in range(0, min(ln, 18) - 1, 2):
                b = bytearray(blob)
                struct.pack_into("<H", b, at + pos + 4 + f, 0xFFFF)
                yield f"biff:{op:#x}+{f}=ffff", bytes(b)
                if f % 4 == 0 and f + 4 <= ln:
                    b = bytearray(blob)
                    struct.pack_into("<I", b, at + pos + 4 + f, 0x7FFFFFFF)
                    yield f"biff:{op:#x}+{f}=7fffffff", bytes(b)
            b = bytearray(blob)
            struct.pack_into("<H", b, at + pos + 2, 0xFFFF)
            yield f"biff:{op:#x}:len=ffff", bytes(b)
        pos += 4 + ln


# ---------------------------------------------------------------------- RTF
def rtf_mutants(name, blob, rng):
    seen = set()
    for m in re.finditer(rb"\\([a-z]+)(-?\d+)", blob):
        if m[1] in seen:
            continue
        seen.add(m[1])
        for hv in (b"2147483647", b"99999999999"):
            yield f"rtf:{m[1].decode()}={hv.decode()}:all", re.sub(rb"(\\" + m[1] + rb")-?\d+", rb"\g<1>" + hv, blob)


FIXTURES = {
    "sheet.xlsx": zip_mutants, "kami_meisai.xlsx": zip_mutants, "sheet.ods": zip_mutants, "form.odt": zip_mutants,
    "form.docx": zip_mutants, "form.hwpx": zip_mutants, "slides.pptx": zip_mutants,
    "acroform.pdf": pdf_mutants, "flat.pdf": pdf_mutants, "scanned.pdf": pdf_mutants, "vertical.pdf": pdf_mutants,
    "rotated.pdf": pdf_mutants, "indic.pdf": pdf_mutants,
    "sheet.xls": cfb_mutants, "form.doc": cfb_mutants, "form.jtd": cfb_mutants,
    "form.rtf": rtf_mutants,
    "bundle.zip": None, "items.csv": None, "items_sjis.csv": None, "form.json": None, "notes.txt": None,
    "encrypted.pdf": None, "rtl.pdf": None,
}


def main():
    out = sys.argv[1]
    os.makedirs(out, exist_ok=True)
    rows = []
    for fx, gen in FIXTURES.items():
        p = os.path.join(PUB, fx)
        if not os.path.exists(p):
            p = os.path.join(ROOT, "tests", "fixtures", fx)
        blob = open(p, "rb").read()
        rng = random.Random(fx)
        items = list(flips(fx, blob, rng))
        if gen:
            items += list(gen(fx, blob, rng))
        seen = set()
        for i, (tag, data) in enumerate(items):
            h = hash(data)
            if h in seen or data == blob:
                continue
            seen.add(h)
            fn = f"{fx.rsplit('.', 1)[0]}__{i:04d}.{fx.rsplit('.', 1)[1]}"
            with open(os.path.join(out, fn), "wb") as f:
                f.write(data)
            rows.append({"file": fn, "fixture": fx, "tag": tag})
        print(fx, len(seen))
    random.Random(7).shuffle(rows)  # 形式を混ぜる（途中でも全形式が見える）
    with open(os.path.join(out, "MANIFEST.jsonl"), "w") as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    print("total", len(rows))


if __name__ == "__main__":
    main()
