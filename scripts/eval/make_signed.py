"""自分で作る、電子署名のある PDF（テキストの欄 1 つ + 署名の欄）。鍵は作るたびに新しく、架空の名前。公開の書類は使わない。
使い方: make_signed.py <出力のフォルダ>  → signed.pdf"""

import subprocess
import sys
from pathlib import Path


def make(out: Path) -> Path:
    out.mkdir(parents=True, exist_ok=True)
    key, crt = out / "k.pem", out / "c.pem"
    subprocess.run(["openssl", "req", "-x509", "-newkey", "rsa:2048", "-nodes", "-keyout", key, "-out", crt, "-subj", "/CN=todoku eval (fictional)", "-days", "30"],
                   check=True, capture_output=True)
    cont = 8192 * 2
    objs = {
        1: b"<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [4 0 R 5 0 R] /SigFlags 3 /NeedAppearances false /DR << /Font << /H 7 0 R >> >> /DA (/H 12 Tf 0 g) >> >>",
        2: b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        3: b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 8 0 R /Resources << /Font << /H 7 0 R >> >> /Annots [4 0 R 5 0 R] >>",
        4: b"<< /Type /Annot /Subtype /Widget /FT /Tx /T (name) /Rect [150 700 450 725] /F 4 /P 3 0 R /DA (/H 12 Tf 0 g) /MK << /BC [0 0 0] >> >>",
        5: b"<< /Type /Annot /Subtype /Widget /FT /Sig /T (Sig1) /Rect [150 600 450 650] /F 4 /P 3 0 R /V 6 0 R >>",
        7: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    }
    stream = b"BT /H 14 Tf 60 705 Td (Name:) Tj ET BT /H 14 Tf 60 620 Td (Signature:) Tj ET"
    objs[8] = b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream"
    objs[6] = (b"<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /Name (test) /M (D:20261105000000Z) "
               b"/ByteRange [0 0000000000 0000000000 0000000000] /Contents <" + b"0" * cont + b"> >>")
    blob = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
    off = {}
    for k in sorted(objs):
        off[k] = len(blob)
        blob += b"%d 0 obj\n" % k + objs[k] + b"\nendobj\n"
    x = len(blob)
    blob += b"xref\n0 9\n0000000000 65535 f \n" + b"".join(b"%010d 00000 n \n" % off[k] for k in range(1, 9))
    blob += b"trailer\n<< /Size 9 /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % x
    i = blob.index(b"/Contents <") + len(b"/Contents ")
    a, b_ = i, i + cont + 2
    rng = b"[0 %010d %010d %010d]" % (a, b_, len(blob) - b_)
    j = blob.index(b"/ByteRange [") + len(b"/ByteRange ")
    blob[j:j + len(rng)] = rng
    dp, sigp = out / "data.bin", out / "sig.der"
    dp.write_bytes(bytes(blob[:a] + blob[b_:]))
    subprocess.run(["openssl", "cms", "-sign", "-binary", "-in", dp, "-signer", crt, "-inkey", key, "-outform", "DER", "-md", "sha256", "-out", sigp],
                   check=True, capture_output=True)
    h = sigp.read_bytes().hex().encode()
    assert len(h) <= cont
    blob[a + 1:a + 1 + len(h)] = h
    p = out / "signed.pdf"
    p.write_bytes(bytes(blob))
    return p


def verify(path: Path, work: Path) -> dict:
    """署名を openssl で確かめる（CMS の検証。証明書の信頼は見ない）。covers_whole：署名がファイルの全部を覆うか。"""
    import re

    b = Path(path).read_bytes()
    m = None
    for m in re.finditer(rb"/ByteRange \[0 (\d+) (\d+) (\d+)\]", b):
        break
    if not m:
        return {"found": False}
    a1, s2, l2 = map(int, m.groups())
    if s2 + l2 > len(b):
        return {"found": True, "valid": False, "why": "範囲がファイルの外"}
    dp, sigp = work / "v_data.bin", work / "v_sig.der"
    dp.write_bytes(b[:a1] + b[s2:s2 + l2])
    hx = b[a1 + 1:s2 - 1].rstrip(b"0")
    if len(hx) % 2:
        hx += b"0"
    try:
        sigp.write_bytes(bytes.fromhex(hx.decode()))
    except ValueError:
        return {"found": True, "valid": False, "why": "署名の字が読めない"}
    r = subprocess.run(["openssl", "cms", "-verify", "-binary", "-inform", "DER", "-in", sigp, "-content", dp, "-noverify", "-out", "/dev/null"],
                       capture_output=True, text=True)
    return {"found": True, "valid": r.returncode == 0, "covers_whole": s2 + l2 == len(b)}


if __name__ == "__main__":
    print(make(Path(sys.argv[1])))
