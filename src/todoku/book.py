"""書類のファイルを開く：形式を中身（先頭のバイト）で見分けて、同じ形の読み取り結果を返す。"""

from __future__ import annotations


class FormatError(ValueError):
    pass


def kind_of(blob: bytes, filename: str = "") -> str:
    """形式の名前。filename（ファイルの名前）は、中身だけでは決まらない形式（CSV・TSV・JSON・JSON Lines）を見分けるのに使う
    （名前を渡さなければ、いままでと同じ判定）。"""
    if blob[:2] == b"PK":
        if len(blob) < 64:
            return "xlsx"
        try:
            from .zipraw import RawZip

            names = set(RawZip(blob).names())
        except Exception:  # noqa: BLE001
            return "xlsx"
        if "word/document.xml" in names or any(n.startswith("word/") and n.endswith("document.xml") for n in names):
            return "docx"
        if any(n.startswith("ppt/") for n in names):
            return "pptx"
        if _is_hwpx(blob, names):
            return "hwpx"  # 韓国のワープロ（한글）の HWPX（OWPML、KS X 6101）：読む・書く（hwpx/）
        odf = _odf_kind(blob, names)  # OpenDocument（.odt・.ods）：mimetype の部品で見分ける
        if odf:
            return odf
        if "[Content_Types].xml" not in names:
            return "zip"  # Office の書類（OOXML：部品の種類の一覧が必ずある）でない ZIP（CSV をまとめたものなど）
        return "xlsx"
    if blob[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        from .cfb import CFB

        names = {n.lower() for n in CFB(blob).streams()}
        if "encryptedpackage" in names:
            return "encrypted"
        if "workbook" in names or "book" in names:
            return "xls"
        if "worddocument" in names:
            return "doc"
        # 一太郎（.jtd・.jtt）：JustSystems のストリームの一覧と、本文（または圧縮した文書）
        from .jtd.book import is_jtd

        if is_jtd(names):
            return "jtd"
        if "fileheader" in names and _is_hwp5(blob):
            return "hwp"  # 韓国のワープロ（한글）の HWP 5.0：FileHeader のストリームが「HWP Document File」で始まる（読むだけ。hwpx/hwp5.py）
        return "cfb"
    if blob[:4] == b"%PDF":
        return "pdf"
    text = _text_kind(blob, filename)  # RTF（中身で）・CSV・TSV・JSON・JSON Lines・文字だけのファイル（.txt）（名前で）
    if text:
        return text
    return "unknown"


# ====================================================================== CSV・JSON・RTF・OpenDocument（中身と名前で見分ける）
_ODF_MIME = {"application/vnd.oasis.opendocument.text": "odt", "application/vnd.oasis.opendocument.text-template": "odt",
             "application/vnd.oasis.opendocument.text-master": "odt",
             "application/vnd.oasis.opendocument.spreadsheet": "ods",
             "application/vnd.oasis.opendocument.spreadsheet-template": "ods",
             "application/vnd.oasis.opendocument.presentation": "odp", "application/vnd.oasis.opendocument.graphics": "odg"}
_BY_NAME = {".csv": "csv", ".tsv": "tsv", ".tab": "tsv", ".json": "json", ".jsonl": "jsonl", ".ndjson": "jsonl",
            ".txt": "txt"}


# 見分けるが、まだ読めない形式（開くと「未対応の形式」と伝え、段階は F）。
# HWP・HWPX は hwpx/ で読める（HWPX は書ける・HWP 5 は読むだけ）ようになったので、いまはない
UNSUPPORTED: dict[str, str] = {}


def unsupported_message(kind: str) -> str:
    return (f"未対応の形式：HWP/HWPX（韓国のワープロ）。このファイルは {UNSUPPORTED.get(kind, kind)} の文書で、"
            "とどくはまだ読めません（中身を推測して読むことはしません）。PDF・Word などに書き出したものがあれば、そちらを開いてください")


def _is_hwpx(blob: bytes, names: set) -> bool:
    """HWPX（OWPML）のパッケージ：先頭の mimetype の部品が application/hwp+zip（ODF と同じ作りのパッケージ）。"""
    if "mimetype" not in names:
        return False
    from .zipraw import RawZip

    try:
        return RawZip(blob).read("mimetype").decode("ascii", "replace").strip() == "application/hwp+zip"
    except Exception:  # noqa: BLE001
        return False


def _is_hwp5(blob: bytes) -> bool:
    """HWP 5.0（複合ファイル）：FileHeader のストリームの先頭 32 バイトが署名「HWP Document File」。"""
    from .cfb import CFB

    try:
        c = CFB(blob)
        name = next(n for n in c.streams() if n.lower() == "fileheader")
        return c.read(name)[:17] == b"HWP Document File"
    except Exception:  # noqa: BLE001
        return False


def _odf_kind(blob: bytes, names: set) -> str | None:
    """OpenDocument のパッケージ（ODF 1.3 第2部 3.3）：先頭の mimetype の部品、なければ manifest.xml の「/」の media-type。"""
    from .zipraw import RawZip

    try:
        z = RawZip(blob)
        if "mimetype" in names:
            mime = z.read("mimetype").decode("ascii", "replace").strip()
        elif "META-INF/manifest.xml" in names:
            import re

            m = re.search(rb'full-path="/"[^>]*?media-type="([^"]+)"|media-type="([^"]+)"[^>]*?full-path="/"',
                          z.read("META-INF/manifest.xml"))
            mime = (m[1] or m[2]).decode("ascii", "replace") if m else ""
        else:
            return None
    except Exception:  # noqa: BLE001
        return None
    return _ODF_MIME.get(mime)


def _text_kind(blob: bytes, filename: str) -> str | None:
    head = blob[:64].lstrip(b"\xef\xbb\xbf \t\r\n")
    if head.startswith(b"{\\rtf"):
        return "rtf"
    if head.startswith(b"<?xml") or head.startswith(b"<office:document"):
        import re

        m = re.search(rb'office:mimetype="([^"]+)"', blob[:4096])  # 1つの XML の OpenDocument（.fodt・.fods）
        if m and _ODF_MIME.get(m[1].decode("ascii", "replace")) in ("odt", "ods"):
            return "f" + _ODF_MIME[m[1].decode("ascii", "replace")]
    name = filename.lower()
    for ext, kind in _BY_NAME.items():
        if name.endswith(ext):
            return kind
    return None


def open_book(blob: bytes, filename: str = ""):
    k = kind_of(blob)
    if k == "xlsx":
        from .xlsx.workbook import Workbook

        return Workbook(blob, filename)
    if k == "xls":
        from .xls.book import XlsWorkbook

        return XlsWorkbook(blob, filename)
    if k == "encrypted":
        raise FormatError("パスワードで暗号化されたファイルです")
    raise FormatError(f"この形式（{k}）は、Excelのブックとして開けません")
