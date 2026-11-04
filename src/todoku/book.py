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
            from .zipraw import AmbiguousZip, RawZip

            names = set(RawZip(blob).names())
        except AmbiguousZip:
            return "zip"  # 中身が二通りに読める ZIP：どの部品で形式を決めるかも決まらない（開くと、その理由を返す）
        except Exception:  # noqa: BLE001
            return "xlsx"
        if "word/document.xml" in names or any(n.startswith("word/") and n.endswith("document.xml") for n in names):
            return "docx"
        if any(n.startswith("ppt/") for n in names):
            return "pptx"
        if "xl/workbook.bin" in names:
            return "xlsb"  # Excel のバイナリのブック（MS-XLSB）：xlsb/
        if _is_hwpx(blob, names):
            return "hwpx"  # 韓国のワープロ（한글）の HWPX（OWPML、KS X 6101）：読む・書く（hwpx/）
        more = _more_zip_kind(blob, names)  # XPS・EPUB・iWork・OpenOffice.org 1.x（ZIP の入れ物のほかの形式）
        if more:
            return more
        odf = _odf_kind(blob, names)  # OpenDocument（.odt・.ods）：mimetype の部品で見分ける
        if odf:
            return odf
        if "[Content_Types].xml" not in names:
            return "zip"  # Office の書類（OOXML：部品の種類の一覧が必ずある）でない ZIP（CSV をまとめたものなど）
        if not any(n.startswith("xl/") for n in names):
            return "ooxml-unknown"  # OOXML のパッケージだが、Excel・Word・PowerPoint の部品がない（Visio・XPS など）
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
        if "powerpoint document" in names:
            return "ppt"  # 古い PowerPoint（MS-PPT）：読むだけ（legacy/ppt.py）
        from .legacy.mail import is_msg

        if is_msg(names):
            return "msg"  # Outlook のメール（MS-OXMSG）
        if "contents" in names and ({"mn0", "spelling", "spelling0", "layout"} & names or filename.lower().endswith(".wps")):
            return "works"  # Microsoft Works（CONTENTS のストリーム）。Kingsoft の .wps は WordDocument があるので doc
        # 一太郎（.jtd・.jtt）：JustSystems のストリームの一覧と、本文（または圧縮した文書）
        from .jtd.book import is_jtd

        if is_jtd(names):
            return "jtd"
        if "fileheader" in names and _is_hwp5(blob):
            return "hwp"  # 韓国のワープロ（한글）の HWP 5.0：FileHeader のストリームが「HWP Document File」で始まる（読むだけ。hwpx/hwp5.py）
        return "cfb"
    if blob[:4] == b"%PDF":
        return "pdf"
    from .imaging.formats import sniff

    if sniff(blob):
        return "image"  # 画像そのもののファイル（JPEG・PNG・GIF・BMP・TIFF・WebP・JPEG 2000・JBIG2・PICT・SVG・HEIC・AVIF）：見て読む（imaging/standalone.py）
    head_kind = _head_kind(blob, filename)  # 先頭のバイトで決まる形式（tar・gzip・LZH・7z・RAR・古い Excel など）
    if head_kind:
        return head_kind
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
            ".eml": "eml", ".txt": "txt", ".html": "markup", ".htm": "markup", ".xhtml": "markup", ".xml": "markup"}


# 見分けるが、まだ読めない形式（開くと「未対応の形式」と伝え、段階は F）。HWP・HWPX は hwpx/ で読む（openkinds）
UNSUPPORTED: dict[str, str] = {
    "ooxml-unknown": "Office Open XML のパッケージ（Excel・Word・PowerPoint の部品がない）",
}


UNSUPPORTED.update({
    "docuworks": "DocuWorks の文書（.xdw・.xbd）",
    "works": "Microsoft Works の文書（.wps）",
    "wk3": "Lotus 1-2-3 の新しい形（.wk3・.wk4・.123）",
})
# 形式ごとの、次の一手（なければ「PDF・Word などに書き出したもの」）
NEXT_STEP: dict[str, str] = {
    "docuworks": "DocuWorks の形は公開されていないので、とどくは読みません（字を推測で取り出しません）。発行元に PDF を頼むか、本人が DocuWorks Viewer Light（無償）で開いて、"
                 "PDF を作るプリンター（Microsoft Print to PDF など）に印刷したものを開いてください",
    "works": "Microsoft Works の形は公開されていないので、とどくは読みません（Kingsoft の .wps は Word と同じ形なので読めます）。"
             "Works か LibreOffice で .docx か PDF に保存し直したものを開いてください",
    "wk3": "とどくは .wk1 までを読みます（.wk3 以降の形には Lotus の公開の資料がないので、推測で読みません）。1-2-3 か LibreOffice で .xlsx か CSV に保存し直したものを開いてください",
}


def _head_kind(blob: bytes, filename: str) -> str | None:
    if blob[:6] == b"7z\xbc\xaf\x27\x1c":
        return "7z"
    if blob[:6] == b"Rar!\x1a\x07":
        return "rar"
    from .legacy.biff234 import is_biff234
    from .legacy.jet import is_jet
    from .legacy.lotus import is_wk1, is_wk3plus
    from .legacy.word2 import is_word2

    if is_biff234(blob):
        return "xls2"  # Excel 2〜4（複合ファイルでない .xls）
    if is_word2(blob):
        return "word2"  # Word 1.x・2.0（複合ファイルでない .doc）
    if is_jet(blob):
        return "mdb"  # Access（Jet・ACE）
    if is_wk1(blob):
        return "wk1"
    if is_wk3plus(blob):
        return "wk3"
    low = filename.lower()
    if low.endswith((".xdw", ".xbd")):
        return "docuworks"  # 中身の形は公開されていない（IANA の登録でも、仕様・印ともに「なし」）。名前で見分ける
    if low.endswith(".dbf"):
        from .legacy.dbf import is_dbf

        if is_dbf(blob):
            return "dbf"
    if low.endswith(".wps"):
        return "works"  # 複合ファイルでない .wps（古い Works）
    from .archives import sniff

    return sniff(blob, filename)


def ooxml_main_type(blob: bytes) -> str | None:
    """OOXML のパッケージの、主な部品の種類（[Content_Types].xml の Override の ContentType。ECMA-376 第2部）：
    パッケージの関係（_rels/.rels）の officeDocument が指す部品、なければ「.main+xml」で終わる種類。分からなければ None。"""
    import re

    from .zipraw import RawZip

    try:
        z = RawZip(blob)
        ct = z.read("[Content_Types].xml").decode("utf-8", "replace")
        rels = z.read("_rels/.rels").decode("utf-8", "replace") if "_rels/.rels" in set(z.names()) else ""
    except Exception:  # noqa: BLE001
        return None
    over = {m[1].lstrip("/"): m[2] for m in re.finditer(r'<(?:\w+:)?Override\b[^>]*?PartName="([^"]*)"[^>]*?ContentType="([^"]*)"', ct)}
    over.update({m[2].lstrip("/"): m[1] for m in re.finditer(r'<(?:\w+:)?Override\b[^>]*?ContentType="([^"]*)"[^>]*?PartName="([^"]*)"', ct)})
    main = re.search(r'Type="[^"]*/officeDocument"[^>]*?Target="([^"]*)"|Target="([^"]*)"[^>]*?Type="[^"]*/officeDocument"', rels)
    if main:
        t = (main[1] or main[2]).lstrip("/")
        if t in over:
            return over[t]
    return next((v for v in over.values() if v.endswith(".main+xml")), None)


def unsupported_message(kind: str, blob: bytes | None = None) -> str:
    what = UNSUPPORTED.get(kind, kind)
    if kind in NEXT_STEP:
        return f"未対応の形式：{what}。{NEXT_STEP[kind]}"
    main = ooxml_main_type(blob) if kind == "ooxml-unknown" and blob else None
    return (f"未対応の形式：{what}" + (f"（主な部品の種類：{main}）" if main else "")
            + "。とどくは読めません（中身を推測して読むことはしません）。PDF・Word などに書き出したものがあれば、そちらを開いてください")


def _more_zip_kind(blob: bytes, names: set) -> str | None:
    """ZIP の入れ物の、ほかの形式：XPS（パッケージの関係 fixedrepresentation）・EPUB（mimetype）。"""
    from .legacy.epub import is_epub
    from .legacy.xps import is_xps
    from .zipraw import RawZip

    try:
        z = RawZip(blob)
    except Exception:  # noqa: BLE001
        return None
    if is_xps(z, names):
        return "xps"
    if is_epub(z, names):
        return "epub"
    from .legacy.iwork import is_iwork
    from .legacy.sxo import kind_of_package

    sx = kind_of_package(z, names)
    if sx:
        return sx
    if is_iwork(names):
        return "iwork"
    return None


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
    k = kind_of(blob, filename)
    if k == "xlsx":
        from .xlsx.workbook import Workbook

        return Workbook(blob, filename)
    if k == "xls":
        from .xls.book import XlsWorkbook

        return XlsWorkbook(blob, filename)
    from .morekinds import WORKBOOKS, open_workbook

    if k in WORKBOOKS:
        return open_workbook(k, blob, filename)
    if k == "encrypted":
        raise FormatError("パスワードで暗号化されたファイルです")
    raise FormatError(f"この形式（{k}）は、Excelのブックとして開けません")
