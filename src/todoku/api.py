"""形式を問わない入口：開く・窓・入力欄・書き込み（証明つき）・検査。コマンドと MCP サーバーの両方が使う。

形式はファイルの中身で見分ける（Excel .xlsx/.xlsm・古い Excel .xls・PDF・Word .docx・古い Word .doc）。
書き込みは、元のファイルを決して上書きしない（別の名前で書き出す）。
"""

from __future__ import annotations

import hashlib
import re
import struct
from pathlib import Path

from .book import UNSUPPORTED, FormatError, kind_of, unsupported_message

KIND_JA = {"xlsx": "Excel", "xls": "古い Excel（.xls）", "pdf": "PDF", "docx": "Word", "doc": "古い Word（.doc）"}
# 一太郎（.jtd）：読む・書く（同じ字数の上書きと、本文の字数を変える書き込み。圧縮した文書は縮め直す）
KIND_JA["jtd"] = "一太郎（.jtd）"

# ---------------------------------------------------------------------- 追加の形式（CSV・TSV・JSON・JSON Lines・RTF・OpenDocument）
# 形式ごとの処理は openkinds が受け持つ。この入口には、それを呼ぶ短い区切りだけを足す
from . import openkinds as _open  # noqa: E402

KIND_JA.update(_open.KIND_JA)
_BOOKS = ("xlsx", "xls", *sorted(_open._more.WORKBOOKS))  # 同じブックの形で読めるもの（表の取り出し・計算の説明）
# ZIP：中のファイルを「a.zip!/中の名前」で指して、その形式のまま読む・書く（ziparc）
KIND_JA["zip"] = "ZIP（ファイルをまとめたもの）"
# ほかの、中にファイルが入っている形式（tar・gzip・LZH・メール）：archives。指定のしかたは ZIP と同じ
from . import archives as _arc  # noqa: E402

KIND_JA.update(_arc.KIND_JA)


class ApiError(ValueError):
    pass


class Unsupported(ApiError):
    """見分けたが、まだ読めない形式（HWP・HWPX）。段階は F（grade は、この形式でも段階を返す）。"""

    def __init__(self, message: str, kind: str):
        super().__init__(message)
        self.kind = kind


class PasswordNeeded(ApiError):
    """開くのにパスワードが要る：本人に聞いて、password に入れてもう一度呼ぶ。"""


class PasswordWrong(ApiError):
    """パスワードが違う：本人にもう一度確かめる。"""


def _read(path: str, password: str | None = None):
    """(読む中身, ファイルの名前, 形式, 暗号の情報)。パスワードで保護された書類は、解いた中身を返す（暗号の情報は、
    書いた後に同じパスワードで保護し直すのに使う。パスワードそのものは結果に出さない）。"""
    from .unlock import NeedsPassword, WrongPassword, unlock

    p = Path(path).expanduser()
    _arc.use_password(password)
    if not p.is_file() and _member(path) is not None:
        return _read_member(path, password)
    if not p.is_file():
        raise ApiError(f"ファイルが見つかりません: {path}")
    blob = _read_file(p)
    try:
        plain, lock = unlock(blob, password)
        _arc.preflight(plain, kind_of(plain, p.name))
    except NeedsPassword as e:
        raise PasswordNeeded(f"{e}。本人にパスワードを聞いて、password に入れてもう一度呼んでください"
                             "（パスワードは、結果にも記録にも残しません）") from None
    except WrongPassword as e:
        raise PasswordWrong(f"{e}。本人にもう一度確かめてください") from None
    except _arc.ArchiveError:
        pass  # 壊れた書庫：窓・段階の所で、理由を添えて伝える
    k = kind_of(plain, p.name)  # 名前は、中身だけでは決まらない形式（CSV・JSON など）を見分けるのに使う
    if k in UNSUPPORTED:
        raise Unsupported(unsupported_message(k, plain), k)  # 見分けたが、まだ読めない形式（book.UNSUPPORTED）
    if k not in KIND_JA:
        raise ApiError(f"この形式（{k}）は扱えません。Excel・PDF・Word のファイルを指定してください")
    return plain, p.name, k, lock


def _read_file(p: Path) -> bytes:
    """ファイルを読む。limits.MAX_FILE より大きなファイルは読まない（申請書類は、ふつう数 MB まで）。"""
    from .limits import MAX_FILE

    size = p.stat().st_size
    if size > MAX_FILE:
        raise ApiError(f"ファイルが大きすぎるので開きません（{size / 1048576:,.0f} MB。開ける大きさは {MAX_FILE / 1048576:,.0f} MB まで。"
                       "環境変数 TODOKU_MAX_FILE_MB で変えられます）")
    return p.read_bytes()


def _budgeted(fn):
    """1 回の呼び出しに、展開・描画の予算を 1 つ置く（limits.call_budget。呼び出しの中で api を呼び直しても同じ予算）。"""
    import functools

    @functools.wraps(fn)
    def run(*a, **k):
        import struct

        from .cfb import CfbError
        from .limits import call_budget
        from .xls.book import XlsError

        try:
            with call_budget():
                return fn(*a, **k)
        except (struct.error, CfbError) as e:
            # 記録の長さ・位置の合わない、壊れたファイル：読み方の内部のエラー（struct.error）をそのまま見せない
            raise ApiError(broken_message(a[0] if a else k.get("path"), e)) from None
        except XlsError as e:
            raise ApiError(str(e)) from None  # 読めないブック：本人に分かる言葉（例外の種類の名前は見せない）

    return run


def broken_message(path, e: Exception) -> str:
    """壊れたファイルを読めなかったときの、本人とエージェントに分かる言い方（複合ファイルなら「壊れた Office の旧形式」）。"""
    import struct

    from .cfb import BROKEN, MAGIC, CfbError

    if isinstance(e, CfbError) and not isinstance(e, struct.error):
        return str(e)
    head = b""
    try:
        p = Path(str(path)).expanduser()
        if p.is_file():
            with p.open("rb") as f:
                head = f.read(8)
    except (OSError, ValueError):
        head = b""
    if head == MAGIC:
        return f"{BROKEN}（.xls・.doc など）：記録の長さや位置が合わず、読めません。元のファイルが壊れていないか、もう一度受け取り直してください"
    return "ファイルが壊れていて読めません（中の記録の長さや位置が合いません）。元のファイルが壊れていないか、もう一度受け取り直してください"


def _member(path: str) -> tuple[str, list[str]] | None:
    """「a.zip!/b.csv」の形の指定（a.zip がファイルとしてあるもの）。"""
    from . import ziparc

    sp = ziparc.split(path)
    if sp is None or not Path(sp[0]).expanduser().is_file():
        return None
    return sp


def _member_blob(path: str) -> tuple[bytes, str]:
    """ZIP の中のファイルの中身と名前（ZIP の中の ZIP も、続けて「!/」で指す）。"""
    from . import ziparc
    from .unlock import NeedsPassword, WrongPassword

    outer, chain = _member(path)
    blob = _read_file(Path(outer).expanduser())
    cur = Path(outer).name
    for m in chain:
        k = kind_of(blob, cur)
        if k not in _arc.KINDS:
            raise ApiError(f"「{m}」の外側のファイルは ZIP ではありません")
        try:
            blob = _arc.extract(blob, m, k, cur)
        except (NeedsPassword, WrongPassword):
            raise  # 7z・RAR のパスワード：呼んだ側（_read_member）が、聞き直す形に直す
        except ziparc.ArchiveError as e:
            raise ApiError(str(e)) from None
        cur = ziparc.base_name(m)
    return blob, ziparc.base_name(chain[-1])


def _read_member(path: str, password: str | None):
    from .unlock import NeedsPassword, WrongPassword, unlock

    try:
        blob, name = _member_blob(path)
        plain, lock = unlock(blob, password)
    except NeedsPassword as e:
        raise PasswordNeeded(f"{e}。本人にパスワードを聞いて、password に入れてもう一度呼んでください"
                             "（パスワードは、結果にも記録にも残しません）") from None
    except WrongPassword as e:
        raise PasswordWrong(f"{e}。本人にもう一度確かめてください") from None
    k = kind_of(plain, name)
    if k in UNSUPPORTED:
        raise Unsupported(f"ZIP の中の「{name}」：" + unsupported_message(k, plain), k)
    if k not in KIND_JA:
        raise ApiError(f"ZIP の中の「{name}」の形式（{k}）は扱えません")
    return plain, name, k, lock


def _encoding(blob: bytes, k: str, encoding: str | None) -> str | None:
    """エージェントが選んだ文字コード（encoding）を確かめる：文字だけのファイル（CSV・TSV・JSON・JSON Lines・.txt）にだけ使え、
    ファイルがその文字コードとして厳密に読めること。読めなければ ApiError（黙ってほかの文字コードで読まない）。"""
    if not encoding:
        return None
    if k not in _open.TEXT_FILES:
        raise ApiError(f"encoding（文字コード）は、文字だけのファイル（CSV・TSV・JSON・JSON Lines・.txt）にだけ指定できます"
                       f"（この書類は {KIND_JA.get(k, k)}。文字コードはファイルの作りで決まります）")
    from .datafile.textenc import EncodingError, decode

    try:
        return decode(blob, encoding).encoding
    except EncodingError as e:
        raise ApiError(str(e)) from None


def _pdf(blob: bytes, name: str, lock):
    from .pdf.doc import PdfDocument
    from .unlock import pdf_password

    pw = pdf_password(lock)
    return PdfDocument(blob, name, password=pw) if pw else PdfDocument(blob, name)


# 窓の控え：対話では、同じファイルを開いてから書く（open_document → fill_document）ので、中身が同じなら作り直さない。
# 鍵はファイルの中身のハッシュと名前だけ（パスワードで保護された書類は控えない。パスワードは控えにも残さない）
_WINDOWS: "dict[tuple, tuple[str, bytes]]" = {}
_WINDOWS_MAX = 8  # 控える書類の数
# 控えの大きさ（窓を写したバイトの合計）の上限：limits.WINDOW_CACHE。大きな書類の窓をいくつも控えて、メモリを使い切らない


@_budgeted
def window(path: str, password: str | None = None, encoding: str | None = None) -> tuple[str, dict, list[tuple[str, str]]]:
    """(形式, 窓の JSON, Markdown の部分)。encoding：文字だけのファイルの文字コードを、エージェントが候補から選んだとき。"""
    import hashlib
    import pickle

    blob, name, k, lock = _read(path, password)
    encoding = _encoding(blob, k, encoding)
    key = (hashlib.sha256(blob).hexdigest(), name, encoding) if lock is None and not (password and k in _arc.PASSWORDED) else None
    if key is not None and key in _WINDOWS:
        _WINDOWS[key] = _WINDOWS.pop(key)  # いちばん新しく使ったものにする
        return pickle.loads(_WINDOWS[key][1])  # 呼んだ側が変えても控えは変わらないよう、写しを返す
    out = _window(blob, name, k, lock, encoding)
    _list_images(out, blob, name, k, lock)
    if key is not None:
        from .limits import WINDOW_CACHE

        try:
            snap = pickle.dumps(out, protocol=pickle.HIGHEST_PROTOCOL)
        except Exception:  # noqa: BLE001  写せないものが入っていれば、控えない
            return out
        if len(snap) <= WINDOW_CACHE:
            _WINDOWS[key] = (k, snap)
        while len(_WINDOWS) > _WINDOWS_MAX or sum(len(v[1]) for v in _WINDOWS.values()) > WINDOW_CACHE:
            _WINDOWS.pop(next(iter(_WINDOWS)))  # いちばん前に使ったものから
    return out


def _list_images(out: tuple, blob: bytes, name: str, k: str, lock) -> None:
    """窓に、書類の中の画像の一覧を足す（doc["images"] と、窓の終わりの節。view_image で見る）。"""
    from .imaging.docimages import list_images, markdown, pdf_annotation_images

    _k, doc, parts = out
    if k == "pdf":
        try:
            images = pdf_annotation_images(_pdf(blob, name, lock))
        except Exception:  # noqa: BLE001  一覧を作れないだけ（窓は出す）
            images = []
    else:
        images = list_images(k, blob, doc)
    if images:
        doc["images"] = images
        parts[-1] = (parts[-1][0], parts[-1][1].rstrip("\n") + "\n" + markdown(images))


def _window(blob: bytes, name: str, k: str, lock, encoding: str | None = None) -> tuple[str, dict, list[tuple[str, str]]]:
    if k in _arc.KINDS:
        from .ziparc import ArchiveError

        try:
            doc, parts = _arc.window(blob, name, KIND_JA, k)
        except ArchiveError as e:
            raise ApiError(str(e)) from None
        return k, doc, parts
    if k in _open.KINDS:  # 追加の形式
        try:
            doc, parts = _open.window(k, blob, name, encoding)
        except _open.OpenFormatError as e:
            raise ApiError(str(e)) from None
        return k, doc, parts
    if k in ("xlsx", "xls"):
        from .book import open_book
        from .xlsx import window as xw

        try:
            wb = open_book(blob, name)
        except FormatError as e:
            raise ApiError(str(e)) from None
        doc = xw.build(wb)
        return k, doc, xw.markdown_parts(doc)
    if k == "pdf":
        from .pdf.window import build, markdown_parts

        doc = build(_pdf(blob, name, lock))
        return k, doc, markdown_parts(doc)
    if k == "jtd":
        # 一太郎：Word と同じ形の窓に、上書きできる字数を添える
        from .jtd.book import JtdError, JtdFile
        from .jtd.window import build as jbuild
        from .jtd.window import markdown_parts as jparts

        try:
            doc = jbuild(JtdFile(blob, name))
        except JtdError as e:
            raise ApiError(f"一太郎の文書として読めません：{e}") from None
        return k, doc, jparts(doc)
    from .docx.window import build, markdown_parts

    d = _word(blob, name, k)
    if getattr(d, "word6", False):
        # Word 6.0/95：上書きできるバイト数（窓では全角の字数）と、字のない欄（変更できない欄）を添える
        from .doc.word6 import build as build6

        doc = build6(d)
    else:
        doc = build(d)
    return k, doc, markdown_parts(doc)


def _is_word6(blob: bytes) -> bool:
    from .doc.word6 import is_word6

    return is_word6(blob)


def _word(blob: bytes, name: str, k: str):
    if k == "doc":
        from .doc.book import DocFile

        return DocFile(blob, name)
    from .docx.document import DocxDocument

    return DocxDocument(blob, name)


@_budgeted
def fields(path: str, password: str | None = None, encoding: str | None = None) -> list[dict]:
    """入力欄の一覧（名前・見出し・種類・選択肢・いまの値・書けるか・その行の文脈）。"""
    k, doc, _ = window(path, password, encoding)
    if k in _arc.KINDS:
        return []  # ZIP そのものに欄はない（中のファイルを「a.zip!/名前」で開く）
    if k in _open.KINDS:  # 追加の形式
        return [_field(f, **w) for f, w in _open.field_rows(k, doc)]
    out = []
    if k in ("xlsx", "xls"):
        for sh in doc["sheets"]:
            for f in sh.get("fields", []):
                out.append(_field(f, sheet=sh["name"]))
    elif k == "pdf":
        for pg in doc["pages"]:
            for f in pg["fields"]:
                out.append(_field(f, page=pg["number"]))
        # XFA のフォーム：ページに置かれていない XFA の欄（動的なフォームの欄など）も、同じ形で出す
        seen = {f["id"] for f in out}
        for f in (doc.get("xfa") or {}).get("fields", []):
            if f["id"] not in seen:
                d = _field(f)
                if f.get("why_not"):
                    d["why"] = f["why_not"]  # 書けない理由
                out.append(d)
    else:
        for f in doc["fields"]:
            out.append(_field(f))
    return out


_SIDE_JA = {"left": ("左", "列"), "right": ("右", "列"), "above": ("上", "行")}


def far_label(f: dict) -> str | None:
    """表計算の欄の見出しが、すぐ隣のマスの字でないとき（離れた所の字）：見出しにせず、番地と、いちばん近い字とその離れ方で
    示す字（「B12（左に 3 列あけて『氏名』）」）。すぐ隣のマスの字が見出しなら None。どれを見出しとして読むかはエージェント。"""
    near = f.get("near")
    lab = f.get("label") or ""
    if not isinstance(near, dict) or not near or not lab:
        return None
    # あいだのマスの数：左右は cols（あいだの列の数）。上は rows（行の番号の差：すぐ上の行は 1）から 1 を引いた、あいだの行の数
    gap = {s: (v.get("cols") if s != "above" else (v["rows"] - 1 if isinstance(v.get("rows"), int) else None))
           for s, v in near.items() if isinstance(v, dict) and s in _SIDE_JA and not v.get("fields_between")}
    # あいだに空の欄をはさむ字（表の格子の見出し・1 桁ずつのマスの見出し）は、離れた字として言い換えない（method に欄の番地）
    if any(isinstance(v, dict) and v.get("fields_between") and (v.get("label") or "") == lab for v in near.values()):
        return None
    if any(g == 0 and (near[s].get("label") or "") and (near[s].get("label") or "") in lab for s, g in gap.items()):
        return None  # すぐ隣のマスの字が見出しにある
    # いちばん近い字（数の値の事実 number は、字として示さない）
    sides = [(near[s].get("mm", 0) or 0, s) for s, g in gap.items() if isinstance(g, int) and not near[s].get("number")]
    if not sides:
        return None
    _mm, s = min(sides)
    ja, unit = _SIDE_JA[s]
    v = near[s]
    if lab not in {str(x.get("label") or x.get("text") or "") for x in near.values() if isinstance(x, dict)}:
        return None  # 見出しが、離れた 1 つの字そのものではない（区分と見出しをつないだ見出し・字にはさまれた欄の並び）
    cell = str(f["id"]).rsplit("!", 1)[-1]
    how = f"{ja}のすぐ隣" if gap[s] == 0 else f"{ja}に {gap[s]} {unit}あけて"
    t = str(v.get("label") or v.get("text") or "")
    t = t if len(t) <= 40 else t[:40] + "…"  # 字のすべては nearest_text と near に（ここは対話で読み上げる短い形）
    return f"{cell}（{how}『{t}』）"


def _field(f: dict, **where) -> dict:
    d = {"id": f["id"], "label": f.get("label", ""), "kind": f.get("kind", ""), "value": f.get("value")}
    far = far_label(f) if where.get("sheet") is not None else None
    if far:
        # 離れた所の字を、欄の見出しとして示さない：番地と、いちばん近い字・その離れ方（近くの字の事実は near）
        d["label"], d["nearest_text"] = far, f.get("label", "")
    for key in ("options", "maxlen", "format", "readonly", "protected", "confidence", "context", "boxes", "write_as",
                "options_depend_on", "unit", "rules", "agent", "why", "placeholder", "cover", "kind_candidates",
                "alternative_to", "group", "why_not", "label_printed", "text_before", "options_labels", "export_value",
                "hidden", "hidden_why", "not_printed", "not_printed_why", "legacy_id", "why_code", "reason_codes",
                "row_label", "column_label", "column_label_partial", "label_clipped", "neighbours", "facts", "near",
                "label_method", "section", "input_reason_codes", "spans", "blanks", "shape", "size", "size_from", "cell_ids",
                "label_key", "referenced_by", "tooltip"):
        if f.get(key) not in (None, "", [], False):
            d[key] = f[key]
    if f.get("agent") and isinstance(f.get("evidence"), dict) and f["evidence"]:
        d["evidence"] = f["evidence"]  # 規則では決められない所の、判断の材料（罫線・塗り・参照している式など）
    for key in ("row_label_gap", "column_label_gap"):  # 数の事実（0 も、すぐ隣という事実）
        if isinstance(f.get(key), int) and not isinstance(f.get(key), bool):
            d[key] = f[key]
    d.update({k: v for k, v in where.items() if v is not None})
    return d


def _insert_spec(key: str, v: dict, k: str) -> dict:
    """字の前・後ろに差し込む指定 {"before" か "after": 字, "value": 書く字, "occurrence"…} を、その形式の書き方に直す。
    PDF は見出しの左・右の空き（after と where）。ほかの形式は、字はそのままで前か後ろに足すだけの old/new（証明も同じ決まりで確かめる）。"""
    if "before" in v and "after" in v:
        raise ApiError(f'{key}: before と after は、どちらか一方を指定してください')
    side = "before" if "before" in v else "after"
    mark = v[side]
    if k == "pdf":
        if side == "after":
            return v
        if "where" in v:
            raise ApiError(f'{key}: before は「字の左」です。where と一緒には指定できません（下は {{"after": 字, "where": "below"}}）')
        return {**{x: y for x, y in v.items() if x != "before"}, "after": mark, "where": "before"}
    sheet = k in ("xlsx", "xls") or k in _open.SHEETS
    allowed = {"occurrence"} | ({"cell", "sheet"} if sheet else {"near", "near_paras"})
    extra = set(v) - {side, "value"} - allowed
    if extra:
        raise ApiError(f"{key}: 分からない指定があります（{'・'.join(sorted(extra))}）。{side} と一緒に使えるのは value・{'・'.join(sorted(allowed))} です")
    if not isinstance(mark, str) or not mark:
        raise ApiError(f'{key}: {side} は、目印にする字（窓の字のとおり）で指定してください')
    val = v.get("value")
    if isinstance(val, bool) or not isinstance(val, (str, int, float)) or val == "":
        raise ApiError(f'{key}: {{"{side}": 目印の字, "value": 書く字}} の形で指定してください（value は字か数）')
    text = _say_value(val)
    return {**{x: y for x, y in v.items() if x in allowed}, "old": mark, "new": text + mark if side == "before" else mark + text}


def _split_values(values: dict, k: str, kinds: dict | None = None) -> tuple[dict, dict]:
    """書く値から、エージェントが述べた意味（as：「生年月日」など）を分ける。返す：(形式ごとの書き手に渡す値, {鍵: 述べた意味})。
    as は対話で伝えるためのもので、書類には書かない。とどくは字の意味を決めないので、as を確かめも直しもしない。
    PDF の kind_candidates のある欄で書く種類を決めるのは kind：{"kind": 種類, "as": 意味, "value": 値}（pdf/fill.py の as_kind）。
    以前の書き方 {"as": 種類, "value": 値} は、kind がなく、as の字が候補とちょうど同じときだけ、種類として受ける。
    {"value": 値} だけになった指定は、値そのものと同じ。@ の字の前・後ろの差し込み（before・after）も、ここで形式の書き方に直す。"""
    kinds = kinds or {}
    out: dict = {}
    said: dict = {}
    for key, v in values.items():
        if isinstance(v, dict):
            v = dict(v)
            if "as" in v and not (k == "pdf" and "kind" not in v and v["as"] in (kinds.get(key) or ())):
                a = v.pop("as")
                if not isinstance(a, str) or not a.strip():
                    raise ApiError(f'{key}: as は、書いた値の意味を字で指定してください（例 {{"value": …, "as": "生年月日"}}）')
                said[key] = a.strip()
            if k == "pdf" and "as" in v:  # 以前の書き方：as が種類の候補と同じ字
                v["kind"] = v.pop("as")
            if isinstance(key, str) and key.startswith("@"):
                if ("before" in v or "after" in v) and not any(x in v for x in ("circle", "seal", "signature", "cover")):
                    v = _insert_spec(key, v, k)
            elif set(v) == {"value"}:
                v = v["value"]
        out[key] = v
    return out, said


@_budgeted
def fill(path: str, values: dict, out_path: str, force: bool = False, password: str | None = None,
         encoding: str | None = None, overwrite: bool = False) -> dict:
    """書き込み、証明し、書き込み後のファイルを読み直して確かめる（_fill）。読み戻しに使った一時ファイルは、必ず消す。"""
    from .inputcheck import InputError, normalize_values

    try:
        # どの形式にも共通の字の確かめ（制御文字・対になっていないサロゲートは書かない。改行は \n にそろえる）
        values = normalize_values(values)
    except InputError as e:
        raise ApiError(str(e)) from None
    temps: list[Path] = []
    try:
        return _fill(path, values, out_path, force, password, encoding, overwrite, temps)
    finally:
        for t in temps:
            t.unlink(missing_ok=True)


def _fill(path: str, values: dict, out_path: str, force: bool, password: str | None, encoding: str | None,
          overwrite: bool, _temps: list) -> dict:
    """書き込み、証明し、書き込み後のファイルを読み直して、書いた値がそのまま出ているかを確かめる。
    パスワードで保護された書類は、解いた中身に書いて証明し、元と同じパスワードで保護し直して書き出す。
    encoding：文字だけのファイルの文字コードを、エージェントが候補から選んだとき（その文字コードで読み、書いた所だけを
    その文字コードのバイトにして置き換える。ほかのバイトは元のまま）。"""
    if _member(path) is not None:
        return _fill_member(path, values, out_path, force, password, encoding, overwrite)
    src = Path(path).expanduser().resolve()
    dst = Path(out_path).expanduser().resolve()
    _not_source(src, dst)
    blob, name, k, lock = _read(path, password)
    encoding = _encoding(blob, k, encoding)
    if k in _arc.KINDS:
        raise ApiError("ZIP そのものには書けません。中のファイルを path に「" + path + "!/中の名前」で指定してください")
    _seal_supported(k, values)
    result: dict = {"kind": k, "out_path": str(dst), "original_sha256": hashlib.sha256(src.read_bytes()).hexdigest()}
    if encoding:
        result["encoding"] = {"encoding": encoding, "chosen_by": "agent",
                              "detail": "エージェントが選んで指定した文字コードで読み、書いた所だけをその文字コードのバイトにしました"}
    readback_values = dict(values)  # 読み戻しで確かめる、欄の id → 書いた値
    told = None  # 本人に伝える、書いた所（ほかの形式は読み戻しの値と同じ）
    rest_fields = None  # まだ空いている欄を数える、欄の一覧（なければ、書いた後のファイルの欄）
    try:
        before = fields(path, password, encoding)
        labels = {f["id"]: f.get("label", "") for f in before}  # 白紙のときの見出し（対話で伝える一覧に使う）
        printed = {f["id"] for f in before if f.get("label_printed")}  # 見出しが文書に印刷される字の欄
        result["_inline"] = {f["id"] for f in before if f.get("kind") == "inline"}  # 文字の中に空欄のあるマス（書く前の窓）
        # 古い数え順の名前（Word と同じ窓の形式の P12.1・P12.a1）：同じ所を指すと決まるもの（legacy_id）だけ、いまの名前に直して書く
        now_ids = {f["id"] for f in before}
        alias = {f["legacy_id"]: f["id"] for f in before if f.get("legacy_id")}
        renamed = {k_: alias[k_] for k_ in values if k_ not in now_ids and k_ in alias and alias[k_] not in values}
        if renamed:
            values = {renamed.get(k_, k_): v for k_, v in values.items()}
            readback_values = dict(values)
            result["renamed_ids"] = renamed
        kind_cands = {f["id"]: f["kind_candidates"] for f in before if f.get("kind_candidates")}
    except ApiError:
        labels, printed, kind_cands = {}, set(), {}
    values, said = _split_values(values, k, kind_cands)
    readback_values = dict(values)
    if k in ("xlsx", "xls"):
        from .book import open_book
        from .xlsx.edit import EditError
        from .xlsx.edit import fill as xfill
        from .xlsx.proof import prove

        circ = {k_: v for k_, v in values.items() if _is_circle(v)}
        cell_values = {k_: v for k_, v in values.items() if k_ not in circ}
        from .xlsx.edit import _group_span_ids
        from .xlsx.window import parse_span_id

        try:
            # マスの文字の中の所の id（「シート!B3@5」）は、セルごとに {位置: 値} にまとめて書く。読み戻しは、書いた後のセルの字で
            cell_values = _group_span_ids(cell_values)
        except EditError as e:
            raise ApiError(str(e)) from None
        readback_values = {k_: v for k_, v in readback_values.items() if k_ not in circ
                           and not (isinstance(k_, str) and "!" in k_ and not k_.startswith("@") and parse_span_id(k_))}
        wb = open_book(blob, name)
        if cell_values or not circ:
            try:
                res = xfill(wb, cell_values, force=force)
            except EditError as e:
                raise ApiError(str(e)) from None
            allowed = res.allowed_inputs
            allowed_c = {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"}  # 部品 → 書いた状態
            pr = prove(blob, res.blob, allowed, allowed_c, {(p.sheet, p.address) for p in res.patches if p.kind == "macro"},
                       rich=set(res.rich_cells))
            new_blob = res.blob
            result.update({"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings),
                           "unexpected": pr.unexpected[:20], "structure_errors": pr.lint_errors[:20],
                           "recalculated": {f"{p.sheet}!{p.address}": p.after for p in res.patches if p.kind == "formula"}})
            # 字のセルの一部を書き換えたセル（空欄ごとの値のリスト・箱の字への true/false・字を目印にした old/new）：
            # 読み戻しと本人に伝える値は、書いた後のセルの字そのもの（true ではなく「■移動支援」）
            for key_, v_ in cell_values.items():
                if isinstance(v_, dict) and "value" in v_:  # {"value": 値, "overwrite_text": true}：読み戻しは値で
                    readback_values[key_] = v_["value"]
            for key_, (id_, text_) in (res.text_written or {}).items():
                readback_values.pop(key_, None)
                readback_values[id_] = text_
            if res.anchored:
                result["anchored"] = dict(res.anchored)
            if res.aligned:
                # 空きの中にそろえて書いた所：そろえ方と、空きの幅を保つために前・後ろに足した半角の空白の字数。
                # 証明は書いたセルだけが変わったこと、読み戻しは書いた後のセルの字（足した空白を含む）で確かめる
                result["aligned"] = dict(res.aligned)
            if res.unverified:
                # 書いた値に連動するのに、todoku が計算し直せない数式（BAHTTEXT・動的配列など）：保存された結果は古いまま
                result["stale_formulas"] = list(res.unverified)
                result["warnings"].append(
                    "書いた値に連動する数式のうち、" + "、".join(res.unverified[:10]) + ("ほか" if len(res.unverified) > 10 else "")
                    + " は todoku では計算し直せません。Excel で開いたときに計算し直されます（開いたときに計算し直す印を付けました）。"
                    "いまファイルに保存されている値は古い（書いた値を反映していない）ので、その値を本人に伝えないでください")
            if res.written_formatting:
                # 書いたセルの書式（取り消し線・白い字・字と塗りが同じ色など）：書いた字が消したように・見えなく印刷される
                result["written_formatting"] = list(res.written_formatting)
                result["warnings"] += [f"{x['id']}: 書いたセルの書式に「{'・'.join(x['formatting'])}」があります"
                                       for x in res.written_formatting]
            if res.fit:
                # 書いた字が、セルの幅・高さに収まらず切れて見える見込み（列の幅・行の高さ・隣のセルの字から見積もった）
                result["fit"] = list(res.fit)
                result["warnings"] += [x["message"] for x in res.fit]
            if res.exclusive:
                result["warnings"] += list(res.exclusive)
            if res.replaced_text:
                # 印刷された字を、明示（overwrite_text・old/new）なしに置き換えた：証明の想定外の変更として示す（提出させない）
                result["proof_ok"] = False
                result["unexpected"] = list(res.replaced_text) + result["unexpected"]
                result["proof"] += f"。印刷された字を置き換えたセルが {len(res.replaced_text)} か所あります（明示されていない置き換え）"
        else:
            new_blob = blob
            result.update({"proof_ok": True, "proof": "", "warnings": [], "unexpected": [], "structure_errors": [], "recalculated": {}})
        if circ:
            new_blob = _circle_excel(new_blob, name, k, circ, result)
    elif k == "pdf":
        from .pdf.fill import FillError
        from .pdf.fill import fill as pfill
        from .pdf.proof import prove

        try:
            res = pfill(_pdf(blob, name, lock), values, force=force)
        except FillError as e:
            raise ApiError(str(e)) from None
        anchored = {p.id: p.anchor for p in res.patches if getattr(p, "anchor", None)}
        pr = _prove_pdf(prove, blob, res.blob, {p.id for p in res.patches}, anchored, lock)
        new_blob = res.blob
        result.update({"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings),
                       "unexpected": pr.unexpected[:20], "structure_errors": pr.lint_errors[:20],
                       "written_as": {p.id: p.value for p in res.patches
                                      if not p.anchor and str(p.value) != str(values.get(p.id))}})
        if res.denied:
            # 作り手の権限の設定が許していない書き込み（断らずに書き、本人に伝える：pdf/fill.py の _permissions）
            result["permissions_denied"] = list(res.denied)
        # 字を組み立てて（シェーピング）描いた欄：証明は、とどく自身の組み立てと照らすだけで、見た目そのものは確かめていない。
        # 書く段階 B として、エージェントに書いた欄の画像を確かめてもらう
        from .pdf.glyphs import shaped as _shaped

        shaped_ids = [{"id": p.id, "page": p.page} for p in res.patches
                      if p.kind == "input" and isinstance(p.value, str) and _shaped(p.value)]
        if shaped_ids:
            result["check_visually"] = {
                "grade": "B", "fields": shaped_ids,
                "why": "字を組み立てて（シェーピング）描いた欄です。字の並びを、とどくの字の組み立てで確かめました（見た目そのものは"
                       "確かめていません）。view_page でその欄のページの画像を見て、字の形・つながり・並びが正しいかを確かめてください"}
        # 値（/V）と見た目の字が違う欄（選択肢の書き出しの値と見せる字・書式で整えた数）：読み戻しは値で、本人に伝えるのは見た目の字で
        pdf_v = {p.id: p.value for p in res.patches if p.kind in ("input", "control") and not p.anchor}
        pdf_shown = {p.id: (p.value, p.shown) for p in res.patches
                     if p.kind in ("input", "calc") and getattr(p, "shown", None) not in (None, p.value)}
        from .pdf.fill import appearance_readback

        pdf_ap_diff = appearance_readback(_pdf(res.blob, name, lock), res.patches)
        xfa_written = {p.id: p.value for p in res.patches if p.kind == "xfa"}
        if xfa_written:
            # XFA の欄：窓に出る値（日付・数は決まった形、選択肢はデータの値）で読み戻す
            readback_values.update(xfa_written)
        if anchored:
            # 見出しの字を目印にした書き込み：書いた欄の id（読み直した窓では、その名前のフォームの欄）
            result["anchored"] = {p.key: p.id for p in res.patches if p.anchor}
            readback_values = {k_: v for k_, v in values.items() if not k_.startswith("@")}
            readback_values.update({p.id: p.value for p in res.patches if p.anchor and p.kind not in ("seal", "circle")})
        circled = [p for p in res.patches if p.kind == "circle"]
        for p in circled:
            if not p.key.startswith("@") and p.circle.get("option"):
                readback_values[p.key] = p.circle["option"]  # 窓の選択肢に書いた所：書いた後の窓では、選んだ選択肢の名前
        if circled:
            # 印刷された字を○で囲んだ所：読み戻しは、書いた後のファイルのページを読み直して、下で確かめる
            result["circled"] = [{"key": p.key, "id": p.id, "page": p.page, "text": p.circle["text"],
                                  "label": p.circle["label"], "context": p.circle["context"],
                                  "around": p.circle.get("around"),  # 囲んだ字の前後（【】が囲んだ字）
                                  # occurrence で指した、字の並びの中の字：長い字の並びの一部を囲んだことを伝える
                                  **({"inside_word": "長い字の並びの中の字を囲みました（occurrence で指定）"}
                                     if p.circle.get("inside_word") else {}),
                                  **({"loose": True} if p.circle.get("loose") else {}),  # ほかの字の縁に触れる（知らせて描いた）
                                  "style": p.circle.get("style", "ellipse"),
                                  "how": ("ページに字を囲む四角の線を描いた" if p.circle.get("style") == "border"
                                          else "ページに楕円を描いた"),
                                  "ellipse_pt": p.circle["ellipse"],
                                  **({"box_mm": p.circle["box_mm"]} if p.circle.get("box_mm") else {}),
                                  **({"border_pt": p.circle["border"]} if p.circle.get("border") else {}),
                                  # 字の読めない所に枠で描いた○：囲んだ字は確かめていない（本人に伝える）
                                  **({"unread": p.circle["unread"], "uncertain": p.circle["unread"]}
                                     if p.circle.get("unread") else {})}
                                 for p in circled]
        rects = {(p.key or p.id): {"id": p.id, **p.rect_mm} for p in res.patches if getattr(p, "rect_mm", None)}
        if rects:
            # とどくが置いた字の枠（ページの左上からの mm）・字の大きさ・そろえ方：view_page で見て確かめ、違えば box・align で書き直す
            result["written_rects"] = rects
        placed = [p for p in res.patches if p.kind == "seal"]
        if placed:
            # 本人の印影・署名の画像（本人が同意したときだけ）：何を、どこに置いたか。読み戻しは、書いた後のファイルで下で確かめる
            result["placed"] = [{"key": p.key, "id": p.id, "page": p.page, "what": p.seal["what"], "say": p.seal["told"],
                                 "box_mm": [round(v / p.seal["mm"], 1) for v in p.seal["rect"]], "rect_pt": p.seal["rect"], "image_sha256": p.seal["sha256"]}
                                for p in placed]
    elif k == "jtd" or (k == "doc" and _is_word6(blob)):
        # 一太郎・Word 6.0/95：同じ字数（バイト数）の上書きは、ファイル全体で変わったのがその字のバイトだけかを証明する。
        # 一太郎の字数を変える書き込み・圧縮した文書は、直すと決めた構造のほかが同じで、直した構造がそれぞれ思ったとおりかを証明する（jtd/proof.py）。
        # 書いた所の印（しおり）を足さないので、読み戻しは証明の中の「読み直した段落の文字」で確かめる
        from .docx.fill import FillError, written_of

        if k == "jtd":
            from .jtd.audit import prove as jprove
            from .jtd.book import JtdFile
            from .jtd.fill import fill as jfill
        else:
            from .doc.word6 import Word6File as JtdFile
            from .doc.word6 import fill as jfill
            from .doc.word6 import prove as jprove

        try:
            res = jfill(JtdFile(blob, name), values, force=force)
        except FillError as e:
            raise ApiError(str(e)) from None
        pr = jprove(blob, res.blob, written_of(res.patches))
        new_blob = res.blob
        result.update({"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings),
                       "unexpected": (pr.unexpected + pr.readback)[:20], "structure_errors": pr.lint_errors[:20]})
        anchored = {p.key: p.id for p in res.patches if p.anchor}
        if anchored:
            result["anchored"] = anchored
        readback_values = {}
        told = {p.id: p.value for p in res.patches}  # 一太郎：本人に伝える、書いた所と値
        result["jtd_read_back"] = pr.readback[:30]
    elif k in _open.KINDS:  # 追加の形式：書き込みと証明は、形式ごとの部品で
        try:
            res_ = _open.fill(k, blob, name, values, force, encoding)
        except _open.OpenFormatError as e:
            raise ApiError(str(e)) from None
        new_blob = res_["blob"]
        result.update(res_["result"])
        readback_values = res_["readback"]
        labels.update(res_.get("labels") or {})
        if res_.get("told") is not None:
            # 書いた所の印を置けない形式（.txt・PowerPoint）：本人に伝えるのは、証明で読み直した書いた値と、元のファイルの残りの欄
            told = res_["told"]
            rest_fields = res_.get("rest")
    else:
        from .docx.fill import FillError

        circ = {k_: v for k_, v in values.items() if _is_circle(v)}
        if circ and k == "doc":
            raise ApiError("古い Word（.doc）には、○の図形を足す書き方をまだ作っていません（図形を足したことを証明できないため）。"
                           "代わりに、字の前に「○」を足す（例：「○有・無」）か、選ばない字を消して書き直す書き方があります。"
                           "どちらにするか本人に確かめてから、字を目印にした書き込み（old・new）で書いてください")
        values = {k_: v for k_, v in values.items() if k_ not in circ}
        readback_values = {k_: v for k_, v in readback_values.items() if k_ not in circ}
        if values or not circ:
            try:
                if k == "doc":
                    from .doc.audit import prove
                    from .doc.fill import fill as wfill
                else:
                    from .docx.fill import fill as wfill
                    from .docx.proof import prove
                res = wfill(_word(blob, name, k), values, force=force)
            except FillError as e:
                raise ApiError(str(e)) from None
            from .docx.fill import written_of

            pr = prove(blob, res.blob, written_of(res.patches))
            new_blob = res.blob
            result.update({"proof_ok": pr.ok, "proof": pr.headline(), "unexpected": pr.unexpected[:20],
                           "structure_errors": pr.lint_errors[:20], "warnings": list(getattr(res, "warnings", None) or [])})
            # 書いた字の書式（読み戻し）：取り消し線・非表示・自動でない字の色など、書いた字が消したように・目立って見える書式
            fmt = _written_formatting(k, res.blob, name, res.patches)
            if fmt:
                result["written_formatting"] = fmt
                result["warnings"] += [f"{x['id']}: 書いた字の書式に「{'・'.join(x['formatting'])}」があります" for x in fmt]
            # 書いたことで段落の行が折り返す見込み（目安）：証明・読み戻しは字しか見ないので、見た目の変化を伝える
            wrapped = _wrap_word(k, blob, res.blob, name, res.patches)
            if wrapped:
                result["wrap"] = wrapped
                result["warnings"] += wrapped
            anchored = {(p.key or p.id): p.id for p in res.patches if getattr(p, "anchor", None)}
            aligned = {p.id: p.align for p in res.patches if getattr(p, "align", None)}
            if aligned:
                # 空きの中にそろえて書いた所：そろえ方と、空きの幅を保つために前・後ろに足した半角の空白の字数（証明で確かめた）。
                # 読み戻しは、書いた値そのもの（足した空白は欄の値に入れない）
                result["aligned"] = aligned
                for id_ in aligned:
                    if isinstance(values.get(id_), dict):
                        readback_values[id_] = values[id_]["value"]
        else:
            new_blob = blob
            anchored = {}
            result.update({"proof_ok": True, "proof": "", "unexpected": [], "structure_errors": []})
        if anchored:
            # 字を目印にした書き込み：書いた所の id（読み直した窓では［P段落.w番号＝値］）
            result["anchored"] = anchored
            readback_values = {k: v for k, v in readback_values.items() if not k.startswith("@")}
            readback_values.update({p.id: p.value for p in res.patches if getattr(p, "anchor", None)})
        # 欄の書式で書いた値（日付の欄の w:dateFormat「yyyy年M月d日」など）：読み戻しは、書式を通した字で比べる
        shaped = {p.id: p.value for p in (res.patches if values or not circ else []) if not getattr(p, "anchor", None) and p.kind == "input"
                  and p.id in readback_values and not isinstance(readback_values[p.id], bool)
                  and str(p.value) != str(readback_values[p.id]) and not _num_equal(p.value, readback_values[p.id])}
        if shaped:
            result["written_as"] = shaped
            readback_values.update(shaped)
        if circ:
            new_blob = _circle_word(new_blob, name, circ, result, original=blob)
    from . import pkgsig

    sig = pkgsig.find(blob)
    if sig is not None and new_blob != blob:
        # パッケージの電子署名：書いた部品のバイト列が変わるので、署名は無効になる。黙って無効にしない（本人に伝えることの先頭に）
        result["warnings"] = [sig.message] + list(result.get("warnings") or [])
        result["digital_signature"] = dict(sig.to_dict(), invalidated=True)
    written_plain = new_blob  # 書いた中身（保護し直す前）
    if lock is not None:
        # 元と同じパスワードで保護し直し、それを同じパスワードで解いて、書いた中身とバイト一致かを確かめてから書き出す
        from .unlock import relock, unlock

        sealed = relock(lock, new_blob)
        if lock.method != "pdf":
            again, _l = unlock(sealed, password)
            if again != new_blob:
                raise ApiError("保護し直したファイルを同じパスワードで解くと、書いた中身と一致しません（書き出していません）")
        new_blob = sealed
        result["password_protected"] = f"元と同じパスワードで保護したまま書き出しました（{lock.label}）"
    result["new_sha256"] = hashlib.sha256(new_blob).hexdigest()
    # 書き出す前に、同じフォルダの一時ファイルに書いて読み戻す（確かめてから、名前を付け替える：_publish）
    tmp = _temp_beside(dst, new_blob)
    _temps.append(tmp)  # 途中で止まっても、fill が消す
    # 読み戻し：書き込み後のファイルの窓に、書いた値が出ているか
    after_list = fields(str(tmp), password, encoding)
    after = {f["id"]: f.get("value") for f in after_list}
    xl_back: dict = {}  # Excel：書いた所の (読み戻しが合うか, 見えている字)。日付・数は、ブックの数え方と表示形式でそろえて比べる
    if k in _open.KINDS:  # 追加の形式：欄の一覧に出ないマス・値も、書いた後のファイルを読み直して照らす
        after.update(_open.read_back(k, written_plain, name, readback_values, encoding))
    if k in _open.SHEETS:
        # OpenDocument の表計算：Excel と同じく、日付・数をそろえて比べ、見えている字（ファイルに保存された字）で伝える
        try:
            from .odf.sheet import OdsWorkbook
            from .xlsx import readback as _xrb

            ods_before, ods_after = OdsWorkbook(blob, name), OdsWorkbook(written_plain, name)
            xl_back = _xrb.check(ods_after, readback_values)
            result["recalculated_cells"] = _xrb.recalculated(
                ods_before, ods_after, [(i.rsplit("!", 1)[0], i.rsplit("!", 1)[1], v) for i, v in (result.get("recalculated") or {}).items()])
        except Exception as e:  # noqa: BLE001  確かめられなかったことを、黙らずに伝える
            xl_back = {}
            result["recalculation_unchecked"] = (f"書いた後の表を読み直せず、見えている字と、数式の計算し直し（recalculated_cells）を"
                                                 f"確かめられませんでした（{type(e).__name__}）。値の読み戻しは、セルの値だけで照らしました")
            result["warnings"] = list(result.get("warnings") or []) + [result["recalculation_unchecked"]]
    if k in ("xlsx", "xls"):
        # 欄の一覧に出ないセル（番地で書いたセル）も、書いた後のブックのセルの値を直接読んで照らす
        from .book import open_book
        from .xlsx.edit import resolve_target

        try:
            wb_after = open_book(written_plain, name)
            for key in readback_values:
                if key in after or key.startswith("@"):
                    continue
                try:
                    sh_, r_, c_ = resolve_target(wb_after, key)
                except Exception:  # noqa: BLE001
                    continue  # 部品の名前など、セルでない指定
                cell = wb_after.sheet(sh_).cells.get((r_, c_))
                after[key] = cell.value if cell is not None else None
            from .xlsx import readback as _xrb

            xl_back = _xrb.check(wb_after, readback_values)
            for key in readback_values:
                if not labels.get(key) and not key.startswith("@"):
                    t_ = _cell_near_text(wb, key)
                    if t_:
                        labels[key] = t_  # 見出しの空いたセル・欄の一覧にないセル：書く前のブックの近くの字（事実）
            if cell_values or not circ:
                result["recalculated_cells"] = _xrb.recalculated(
                    wb, wb_after, [(p.sheet, p.address, p.after) for p in res.patches if p.kind == "formula"])
        except Exception as e:  # noqa: BLE001  確かめられなかったことを、黙らずに伝える
            result["recalculation_unchecked"] = (f"書いた後のブックを読み直せず、見えている字と、数式の計算し直し（recalculated_cells）を"
                                                 f"確かめられませんでした（{type(e).__name__}）")
            result["warnings"] = list(result.get("warnings") or []) + [result["recalculation_unchecked"]]
    if k == "pdf":
        readback_values.update({i: v for i, v in pdf_v.items() if i in readback_values})
    mismatch = []
    # 読み戻しのしかた：書いた所の印を置けない形式（一太郎・Word 6.0/95・PowerPoint・.txt）は、証明の中で書いた後のファイルを
    # 読み直して照らす（readback_values は空）。ほかは、書いた後のファイルの窓（after）で照らす
    result["read_back_by"] = "proof" if told is not None else "window"
    for key, v in readback_values.items():
        g = after.get(key)
        if key not in after:
            # 書いた所が、書いた後のファイルの窓に見つからない：照らせていないので、合ったことにしない
            mismatch.append({"id": key, "written": v, "read_back": None, "why": "読み戻せなかった（書いた後のファイルに、この欄が見つかりません）"})
            continue
        if key in xl_back:
            ok = xl_back[key][0]
            g = xl_back[key][1]
        elif isinstance(v, bool) or isinstance(g, bool):
            # 書いた値の「はい／いいえ」は、その形式の書き手と同じ読み方で（有・☑・ON・レ・○ など）
            want = _writer_bool(k, v)
            ok = want is not None and bool(g) == want
        else:
            ok = str(g) == str(v) or _num_equal(g, v)
        if not ok:
            mismatch.append({"id": key, "written": v, "read_back": g})
    if k == "pdf":
        mismatch += pdf_ap_diff  # 見た目（/AP）の字：描ける字形で、見せるはずの字か
        for f_ in after_list:
            if f_["id"] in pdf_shown and str(f_.get("value")) == str(pdf_shown[f_["id"]][0]):
                f_["value"] = pdf_shown[f_["id"]][1]  # 本人に伝えるのは、書類に見える字（選択肢の名前・書式で整えた数）
    if result.get("placed"):
        mismatch += _read_back_placed(k, written_plain, name, lock, values, result["placed"])
    if result.get("circled"):
        mismatch += _read_back_circled(k, written_plain, name, lock, result["circled"])
    if "jtd_read_back" in result:
        mismatch += [{"id": "段落", "written": "", "read_back": x} for x in result.pop("jtd_read_back", [])]
    if "text_read_back" in result:  # .txt：証明の中で、書いた後のファイルを読み直した行
        mismatch += [{"id": "行", "written": "", "read_back": x} for x in result.pop("text_read_back", [])]
    if "para_read_back" in result:  # PowerPoint：証明の中で、書いた後のファイルを読み直した段落
        mismatch += [{"id": "段落", "written": "", "read_back": x} for x in result.pop("para_read_back", [])]
    result["read_back_ok"] = not mismatch
    result["read_back_differences"] = mismatch[:30]
    hand = _HAND_ONLY.get(result["original_sha256"])
    if hand is not None:
        # 段階を決めたとき（grade_document の findings）に見つかった、手で行う所（書く D）：本人に伝えることに必ず入れる
        result["hand_only"] = hand
    else:
        result["hand_only_unchecked"] = ("署名・押印など手で行う所は、まだ確かめていません（grade_document で窓を読んで findings に"
                                         "渡すと、書き込みの結果の本人に伝えることに入ります）")
    rects: dict = {}
    if k == "pdf":
        # PDF：同じ所の別の読み方（規則の欄と「？」の所）を、書いた所と重なる欄として、まだ空いている欄から外すのに使う
        try:
            _k, wdoc, _p = window(str(tmp), password)
            for pg in wdoc.get("pages", []):
                for f_ in pg.get("fields", []):
                    if f_.get("rect"):
                        rects[f_["id"]] = (pg.get("number"), f_["rect"])
        except Exception:  # noqa: BLE001  重なりで外せないだけ（欄はそのまま数える）
            rects = {}
    if told is not None:
        # 一太郎：書いた所の印がないので、読み直した窓の同じ id（残りの空白）ではなく、証明で確かめた書いた値を伝える
        rest = rest_fields if rest_fields is not None else [f for f in after_list if f["id"] not in told]
        result.update(_to_tell(rest, told, result, labels, printed, rects, said))
    elif xl_back:
        # Excel：本人に伝える値は、セルの表示形式を通した見えている字（日付は「2026/11/3」、数は「3」）
        shown_list = [dict(f, value=xl_back[f["id"]][1]) if f["id"] in xl_back else f for f in after_list]
        result.update(_to_tell(shown_list, {k_: xl_back[k_][1] if k_ in xl_back else v for k_, v in readback_values.items()},
                               result, labels, printed, rects, said))
    else:
        result.update(_to_tell(after_list, readback_values, result, labels, printed, rects, said))
    _publish(tmp, dst, src, overwrite, result)
    return result


def _cell_near_text(wb, key: str) -> str:
    """セルのいちばん近い字：同じ行を左へ、なければ同じ列を上へたどって、最初にある字（字のまま。見出しと決めたものではない）。"""
    from .xlsx.edit import resolve_target
    from .xlsx.window import label_text

    try:
        sheet, r, c = resolve_target(wb, key)
        cells = wb.sheet(sheet).cells
    except Exception:  # noqa: BLE001  部品の名前など、セルでない指定
        return ""
    for rr, cc in [(r, x) for x in range(c - 1, 0, -1)] + [(x, c) for x in range(r - 1, 0, -1)]:
        cell = cells.get((rr, cc))
        if cell is not None and isinstance(cell.value, str) and not cell.has_formula and cell.value.strip():
            return label_text(cell.value)
    return ""


def _writer_bool(k: str, v) -> bool | None:
    """読み戻しで、書いた値を「はい／いいえ」として読む：その形式の書き手が使うのと同じ読み方（読めなければ None）。"""
    if isinstance(v, bool):
        return v
    try:
        if k in ("xlsx", "xls") or k in _open.SHEETS:
            from .xlsx.edit import as_bool

            return as_bool(v)
        if k == "pdf":
            from .pdf.fill import _as_bool
        else:
            from .docx.fill import _as_bool
        return _as_bool(v)
    except Exception:  # noqa: BLE001  書き手が「はい／いいえ」と読まない値
        return None


def _wrap_word(k: str, before: bytes, after: bytes, name: str, patches=()) -> list[str]:
    """Word（.docx・.doc）：書いたことで段落の行が折り返す見込み（目安）の警告。見積もれなかったことも黙らずに伝える。"""
    try:
        if k == "doc":
            from .doc.wrap import wrap_warnings
        else:
            from .docx.wrap import wrap_warnings
        by_mark = {str(p.id).replace(".", "_").replace("#", "_"): p.id for p in patches}
        return wrap_warnings(_word(before, name, k), _word(after, name, k), by_mark)
    except Exception as e:  # noqa: BLE001
        return [f"書いたことで行が折り返すかを見積もれませんでした（{type(e).__name__}）"]


def _written_formatting(k: str, blob: bytes, name: str, patches) -> list[dict]:
    """Word（.docx・.doc）：書いた後のファイルを読み直し、書いた字の書式に、取り消し線・非表示・自動でない字の色・蛍光ペン・
    網かけがあれば [{id, formatting}]（書いた字が、消したように・目立って印刷されるのを、読み戻しで伝える）。"""
    try:
        if k == "doc":
            from .doc.fill import written_formatting
        else:
            from .docx.fill import written_formatting
        found = written_formatting(blob, name)
    except Exception as e:  # noqa: BLE001  確かめられなかったことを、黙らずに伝える
        return [{"id": "書いた所", "formatting": [f"書いた字の書式を読み直せませんでした（{type(e).__name__}）"]}]
    by_mark = {str(p.id).replace(".", "_").replace("#", "_"): p.id for p in patches}
    return [dict(x, id=by_mark.get(x["id"], x["id"])) for x in found]


def _not_source(src: Path, dst: Path) -> None:
    """書き出す先が、元のファイル（そのものか、同じ中身を指すハードリンク・シンボリックリンク）でないこと。"""
    import os

    if src == dst or (dst.exists() and src.exists() and os.path.samefile(src, dst)):
        raise ApiError("元のファイルは上書きしません。別の名前（out_path）を指定してください")


def unverified_path(dst: Path) -> Path:
    """証明か読み戻しで確かめられなかった書き込みを置く名前：「名前.unverified.拡張子」。"""
    return dst.with_name(f"{dst.stem}.unverified{dst.suffix}")


def _temp_beside(dst: Path, blob: bytes) -> Path:
    """dst と同じフォルダの一時ファイル（拡張子は dst と同じ：中身で決まらない形式も、名前で見分けられる）に書く。"""
    import os
    import tempfile

    dst.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix=f".{dst.stem}.", suffix=f".todoku-tmp{dst.suffix}", dir=dst.parent)
    with os.fdopen(fd, "wb") as f:
        f.write(blob)
        f.flush()
        os.fsync(f.fileno())
    return Path(name)


def _publish(tmp: Path, dst: Path, src: Path, overwrite: bool, result: dict, ok: bool | None = None) -> Path:
    """確かめた一時ファイルを、書き出す名前に付け替える（os.replace：途中で止まっても、半分だけ書いたファイルは残らない）。
    証明か読み戻しで確かめられなかったときは、dst ではなく「名前.unverified.拡張子」に置き、提出してはいけないことを伝える。
    dst にすでに別の中身のファイルがあれば、overwrite のときだけ置き換える（元のファイルとそのリンクは、決して置き換えない）。"""
    import os

    if ok is None:
        ok = bool(result.get("proof_ok")) and bool(result.get("read_back_ok"))
    final = dst if ok else unverified_path(dst)
    _not_source(src, final)
    if ok and final.exists() and not overwrite and final.read_bytes() != tmp.read_bytes():
        raise ApiError(f"書き出す先「{final.name}」には、すでに別のファイルがあります（書き出していません）。"
                       "置き換えてよいと本人が言ったときだけ overwrite に true を入れるか、別の名前（out_path）を指定してください")
    os.replace(tmp, final)
    result["out_path"] = str(final)
    result["verified"] = ok
    if not ok:
        note = (f"この書き込みは、証明か読み戻しで確かめられなかったので、「{final.name}」という名前で書き出しました"
                f"（「{dst.name}」には書いていません）。このファイルは提出しないでください。確認が必要な点を直して、もう一度書き込んでください。")
        result["unverified"] = note
        result["say_to_user"] = note + ("\n" + result["say_to_user"] if result.get("say_to_user") else "")
    return final


def _fill_member(path: str, values: dict, out_path: str, force: bool, password: str | None,
                 encoding: str | None = None, overwrite: bool = False) -> dict:
    """ZIP の中のファイルに書く：取り出したファイルに書いて証明し、ZIP はそのファイルの部品だけを入れ替えて書き出す。
    書き出した ZIP から取り出し直したファイルが、証明した書き込み後のファイルとバイト一致し、ほかの部品が元とバイト一致するかを確かめる。"""
    import tempfile

    from . import ziparc

    outer, chain = _member(path)
    src = Path(outer).expanduser().resolve()
    dst = Path(out_path).expanduser().resolve()
    _not_source(src, dst)
    original = _read_file(src)
    layers = [(original, kind_of(original, src.name), src.name)]  # 外側から順の入れ物：(中身, 形式, 名前)
    try:
        for m in chain[:-1]:
            sub = _arc.extract(layers[-1][0], m, layers[-1][1], layers[-1][2])
            layers.append((sub, kind_of(sub, ziparc.base_name(m)), ziparc.base_name(m)))
    except ziparc.ArchiveError as e:
        raise ApiError(str(e)) from None
    for _b, lk, _n in layers:
        if lk not in _arc.WRITABLE:
            raise ApiError(_arc.NO_WRITE.get(lk, f"この形式（{lk}）の中のファイルには書けません"))
    member_blob, name = _member_blob(path)
    name = name.replace("\\", "_")
    if name.strip(".") == "":  # 「..」「.」だけの名前：一時フォルダの中のファイルの名前にできない
        name = "member"
    with tempfile.TemporaryDirectory() as d:
        a, b = Path(d) / "a" / name, Path(d) / "b" / name
        a.parent.mkdir()
        b.parent.mkdir()
        a.write_bytes(member_blob)
        result = fill(str(a), values, str(b), force, password, encoding)
        written = Path(result["out_path"]).read_bytes()  # 確かめられなければ「名前.unverified.拡張子」
    # 内側から順に、書いた部品だけを入れ替える
    new = written
    for (zblob, lk, ln), m in reversed(list(zip(layers, chain))):
        try:
            replaced = _arc.replace(zblob, m, new, lk, ln)
        except ziparc.ArchiveError as e:
            raise ApiError(str(e)) from None
        problems = _arc.unchanged_except(zblob, replaced, m, lk, ln)
        if problems:
            raise ApiError("ZIP に入れ直すと、書いたファイルのほかの部品も変わります（書き出していません）：" + "、".join(problems[:5]))
        new = replaced
    inner = ziparc.SEP.join(chain)
    tmp = _temp_beside(dst, new)
    try:
        back, _n = _member_blob(str(tmp) + ziparc.SEP + inner)
        if back != written:
            raise ApiError("書き出した ZIP から取り出したファイルが、書き込んで証明したファイルと一致しません（書き出していません）")
        if result.get("unverified"):  # 中のファイル（一時フォルダ）の名前でなく、ZIP の名前で伝え直す（_publish）
            note = result.pop("unverified")
            result["say_to_user"] = (result.get("say_to_user") or "").replace(note + "\n", "").replace(note, "")
        dst = _publish(tmp, dst, src, overwrite, result)
    finally:
        tmp.unlink(missing_ok=True)
    result.update({"out_path": str(dst), "original_sha256": hashlib.sha256(original).hexdigest(),
                   "new_sha256": hashlib.sha256(new).hexdigest(), "member": inner,
                   "archive": "ZIP の中のファイルに書き、ZIP はそのファイルの部品だけを入れ替えました（ほかの部品は元とバイト一致）。"
                              f"書いたファイルは「{dst}{ziparc.SEP}{inner}」で開けます"})
    return result


def _is_circle(v) -> bool:
    """○で囲む指定（{"circle": 囲む字, …}）か。"""
    return isinstance(v, dict) and "circle" in v


def _circle_excel(blob: bytes, name: str, k: str, circ: dict, result: dict) -> bytes:
    """Excel：セルの字を○の図形で囲む。書いた後のブックで○を決め、図形を足し、足しただけであることを証明する。"""
    from .book import open_book
    from .xlsx import circle as xc

    if k == "xls":
        raise ApiError("古い Excel（.xls）には、○の図形を足す書き方をまだ作っていません（図形を足したことを証明できないため）。"
                       "代わりに、字の前に「○」を足す（例：「○有・無」）か、選ばない字を消して書き直す書き方があります。"
                       "どちらにするか本人に確かめてから、そのセルに字として書いてください")
    wb = open_book(blob, name)
    used: dict = {}
    sds = []
    problems = []
    for key, spec in circ.items():
        try:
            sds.append((key, xc.resolve(wb, spec, used)))
        except xc.CircleError as e:
            problems.append(f"{key}: {e}")
    if problems:
        raise ApiError("\n".join(problems) + "\n（代わりに、字の前に「○」を足す・選ばない字を消して書き直す書き方があります。"
                       "本人に確かめてから、そのセルに字として書いてください）")
    try:
        out = xc.apply(open_book(blob, name), [sd for _k, sd in sds])
    except xc.CircleError as e:
        raise ApiError(str(e)) from None
    bad = xc.prove(blob, out, name, [(sd["id"], circ[key_]) for key_, sd in sds])
    result["unexpected"] = list(result.get("unexpected") or []) + bad
    result["proof_ok"] = bool(result.get("proof_ok", True)) and not bad
    result["proof"] = (result.get("proof") or "") + (f"○ {len(sds)} つは、シートの図形に楕円を足しただけです（ほかの部品はバイト一致）"
                                                     if not bad else f"○の書き込みで想定外の変更 {len(bad)} 件")
    result["circled"] = [{"key": key_, "id": sd["id"], "sheet": sd["sheet"], "cell": sd["cell"], "text": sd["text"],
                          "label": sd["label"], "context": sd["context"], "around": sd.get("around"), "style": sd.get("style", "ellipse"),
                          "how": ("シートの図形に四角を足した（○ではなく□。Excel にはセルの字の一部を囲む罫線がないため、図形で囲んだ）"
                                  if sd.get("style") == "border" else "シートの図形に楕円を足した"),
                          "ellipse_pt": sd["ellipse"], "error_mm": sd.get("error_mm"),
                          "uncertain": "位置は見積もり。印刷か画面で確かめが要る" + (f"：{'・'.join(sd['notes'])}" if sd["notes"] else "")}
                         for key_, sd in sds]
    w = list(result.get("warnings") or [])
    w += _both_circled(wb, [sd for _k, sd in sds])
    if xc.UNCERTAIN not in w:
        w.append(xc.UNCERTAIN)
    result["warnings"] = w
    result["anchored"] = {**(result.get("anchored") or {}), **{key_: sd["id"] for key_, sd in sds}}
    return out


def _both_circled(wb, sds: list[dict]) -> list[str]:
    """同じセルの字の、選択肢の全部（セルの字から囲んだ字を除くと、区切りの記号と空白だけが残る：「有・無」「男　女」）を○で
    囲んだ：どちらか1つを選ぶ欄のことが多いので、警告の言葉（書くのは止めない）。"""
    import unicodedata

    by_cell: dict = {}
    for sd in sds:
        by_cell.setdefault((sd["sheet"], sd["row"], sd["col"]), []).append(sd["text"])
    out = []
    for (sheet, r, c), texts in by_cell.items():
        if len(set(texts)) < 2:
            continue
        cell = wb.sheet(sheet).cells.get((r, c))
        rest = cell.value if cell is not None and isinstance(cell.value, str) else None
        if rest is None:
            continue
        for t in set(texts):
            rest = rest.replace(t, "", 1)
        if all(unicodedata.category(ch)[0] in "PZS" for ch in rest):
            from .formula.parser import addr as _addr

            out.append(f"{sheet}!{_addr(r, c)}: 「{cell.value.strip()}」の選択肢の全部（{'・'.join(dict.fromkeys(texts))}）を○で囲みました。"
                       "どちらか1つを選ぶ欄なら、本人に確かめて1つだけにしてください")
    return out


def _circle_word(blob: bytes, name: str, circ: dict, result: dict, original: bytes | None = None) -> bytes:
    """Word（.docx）：本文の字を○で囲む（字に結び付けた楕円の図形、または字の囲み線）。足しただけであることを証明する。
    original：同じ書き込みで値を書く前の文書。○の所はその字で決める（ほかの値を書いたかで、囲む字が変わらないように）。"""
    from .docx import circle as dc
    from .docx.document import DocxDocument

    doc = DocxDocument(blob, name)
    orig = DocxDocument(original, name) if original is not None and original != blob else None
    used: dict = {}
    sds, problems = [], []
    for key, spec in circ.items():
        try:
            sds.append((key, dc.resolve(doc, spec, used, original=orig)))
        except dc.CircleError as e:
            problems.append(f"{key}: {e}")
    if problems:
        raise ApiError("\n".join(problems))
    out = dc.apply(doc, [sd for _k, sd in sds])
    bad = dc.prove(blob, out, name, [(sd["id"], circ[key_]) for key_, sd in sds], original=original)
    result["unexpected"] = list(result.get("unexpected") or []) + bad
    result["proof_ok"] = bool(result.get("proof_ok", True)) and not bad
    result["proof"] = (result.get("proof") or "") + (f"○ {len(sds)} つは、字の前に図形（または字の囲み線）を足しただけです"
                                                     "（ほかの部品はバイト一致）" if not bad else f"○の書き込みで想定外の変更 {len(bad)} 件")
    result["circled"] = [{"key": key_, "id": sd["id"], "text": sd["text"], "label": sd["label"], "context": sd["context"], "around": sd.get("around"),
                          "style": sd["style"], "size": sd["size"],
                          "how": ("字に結び付けた楕円の図形を足した（style: ellipse。位置は見積もり）" if sd["style"] == "ellipse"
                                  else "字に四角い囲み線（w:bdr）を付けた（○ではなく□。字の書式なので位置はずれない）"),
                          "uncertain": ("位置は見積もり。印刷か画面で確かめが要る" + (f"：{'・'.join(sd['notes'])}" if sd.get("notes") else "")
                                        if sd["style"] == "ellipse" else "・".join(sd.get("notes") or []))}
                         for key_, sd in sds]
    if any(sd["style"] == "ellipse" for _k, sd in sds):
        w = list(result.get("warnings") or [])
        if dc.UNCERTAIN not in w:
            w.append(dc.UNCERTAIN)
        result["warnings"] = w
    # 字の囲み線は字の幅を広げるので、付けただけで行が折り返すことがある（値を書いたときと同じに伝える）
    wrapped = _wrap_word("docx", blob, out, name)
    if wrapped:
        result["wrap"] = list(result.get("wrap") or []) + wrapped
        result["warnings"] = list(result.get("warnings") or []) + wrapped
    result["anchored"] = {**(result.get("anchored") or {}), **{key_: sd["id"] for key_, sd in sds}}
    return out


def _read_back_circled(k: str, blob: bytes, name: str, lock, circled: list[dict]) -> list[dict]:
    """○で囲んだ所の読み戻し：書いた後のファイルを開き直し、○がその字だけを囲んでいるか（ほかの字に触れていないか）。"""
    if k == "pdf":
        from .pdf import circle

        return circle.read_back(_pdf(blob, name, lock), circled)
    if k in ("xlsx", "xls"):
        from .xlsx import circle as xcircle

        return xcircle.read_back(blob, name, circled)
    if k in ("docx",):
        from .docx import circle as dcircle

        return dcircle.read_back(blob, name, circled)
    return []


def _seal_supported(k: str, values: dict) -> None:
    """本人の印影・署名の画像（seal・signature）を置けるのは、いまは PDF だけ。ほかの形式は、証明つきで置く方法を
    作るまで、書かずに断る（証明のない書き込みはしない）。"""
    from .sealimage import KINDS, SealError, seal_key

    if k == "pdf":
        return
    for key, v in values.items():
        try:
            sk = seal_key(v)
        except SealError as e:
            raise ApiError(f"{key}: {e}") from None
        if sk is not None:
            raise ApiError(f"{key}: 本人の{KINDS[sk]}の画像を置けるのは、いまは PDF の書類だけです（{KIND_JA.get(k, k)}には、"
                           "変わったのが置いた画像だけであることを証明して置く方法を、まだ作っていません）。"
                           "書いていません。PDF の様式があればそれを使うか、このことを本人に伝えてください")


def _read_back_placed(k: str, blob: bytes, name: str, lock, values: dict, placed: list[dict]) -> list[dict]:
    """置いた画像の読み戻し：書いた後のファイルを開き直し、その名前の画像が、伝えた位置に、本人の画像のとおりにあるか。"""
    from .sealimage import SealError, decode

    out = []
    if k != "pdf":
        return out
    from .pdf import seal

    doc = _pdf(blob, name, lock)
    for pl in placed:
        spec = values.get(pl["key"]) or {}
        key = "seal" if "seal" in spec else "signature"
        try:
            img = decode(Path(str(spec.get(key))).expanduser().read_bytes())
        except (SealError, OSError) as e:
            out.append({"id": pl["id"], "written": pl["what"], "read_back": f"画像を読み直せません：{e}"})
            continue
        # 位置は、置いたときに決めた矩形（見た目の座標、pt）。ファイルの注釈の位置と照らす
        probs, _nums = seal.verify(doc, pl["page"], {"id": pl["id"], "rect": pl["rect_pt"]}, key, img)
        if img.sha256 != pl["image_sha256"]:
            probs.append("画像のファイルが、置いたときと変わっています")
        if probs:
            out.append({"id": pl["id"], "written": pl["what"], "read_back": "／".join(probs)[:300]})
    return out


def _prove_pdf(prove, original: bytes, new: bytes, allowed: set, anchored: dict, lock):
    """PDF の証明。パスワードで保護された PDF は、同じパスワードで開いて確かめる。"""
    import inspect

    from .unlock import pdf_password

    pw = pdf_password(lock)
    if pw and "password" in inspect.signature(prove).parameters:
        return prove(original, new, allowed, anchored, password=pw)
    return prove(original, new, allowed, anchored)


def _say_value(v) -> str:
    """本人に伝える値：数は Excel の標準の表示と同じ字（3.0 ではなく 3）。"""
    if isinstance(v, float):
        from .formula.values import number_to_text

        return number_to_text(v)
    return str(v)


def _overlaps(a: tuple, b: tuple) -> bool:
    """(ページ, [x0, y0, x1, y1]) の2つが、小さいほうの面積の半分以上で重なるか。"""
    if a[0] != b[0]:
        return False
    (ax0, ay0, ax1, ay1), (bx0, by0, bx1, by1) = ([min(r[0], r[2]), min(r[1], r[3]), max(r[0], r[2]), max(r[1], r[3])]
                                                  for r in (a[1], b[1]))
    w, h = min(ax1, bx1) - max(ax0, bx0), min(ay1, by1) - max(ay0, by0)
    small = min((ax1 - ax0) * (ay1 - ay0), (bx1 - bx0) * (by1 - by0))
    return w > 0 and h > 0 and small > 0 and w * h >= 0.5 * small


_DATA_KINDS = ("csv", "tsv", "json", "jsonl")  # 列の名前・キーが、その値の名前（近くの字ではない）


def _verbatim_label(f: dict) -> str:
    """本人に伝える見出し：窓の見出しが「推定した見出し：その行の字（［▶］入り）」の形で、推定した見出しの語がその行（context。
    窓の読み順の字のとおり）にないときは、その語を外し、行の字のとおりにする（ほかの行・列から推定した語「個人番号：」が、
    手帳番号の欄に付いて伝わらないように）。"""
    lab, ctx = f.get("label") or "", f.get("context") or ""
    i = lab.find("［▶")
    p = lab.rfind("：", 0, i) if i > 0 else -1
    if p <= 0 or "［" in lab[:p]:
        return lab

    def flat(s: str) -> str:
        return re.sub(r"\s+", "", s)

    keep = [x for x in lab[:p].split("／") if x and flat(x) in flat(ctx)]
    return "／".join(keep) + "：" + lab[p + 1:] if keep else lab[p + 1:]


def _to_tell(after_list: list[dict], written: dict, result: dict, labels: dict | None = None,
             printed: set | None = None, rects: dict | None = None, said: dict | None = None) -> dict:
    """本人が Word などで開いて確かめなくても、エージェントが対話で伝えられるように：書いた所（見出しと、書いた後の
    ファイルを読み直した値）・まだ空いている欄・本人に伝えること。
    見出しは、書く前の書類の見出し（labels）を使う：書いた後の書類では、いま書いた値が隣の欄の見出しに見えることがある
    （「電話／東京都千代田区1-1」）。見出しは書類に印刷された字で、todoku が書いた値ではない。
    printed：見出しが書類に印刷された字の欄（欄の名前と同じ字でも「見出しなし」にしない）。
    said：エージェントが述べた、書いた値の意味（as）。近くの字は事実であって欄の意味ではないので、欄の名前として示さない：
    as があれば「生年月日：昭和60年4月1日（p1.t45-112／近くの字「住所」）」、なければ「p1.t45-112（近くの字「住所」）：昭和60年4月1日」。"""
    labels = labels or {}
    printed = printed or set()
    after_list = [dict(f, label=labels[f["id"]]) if labels.get(f["id"]) else f for f in after_list]
    after_list = [dict(f, label=_verbatim_label(f)) if f.get("label") and f.get("context") else f for f in after_list]
    by_id = {f["id"]: f for f in after_list}
    told = []
    sheet_kind = result.get("kind") in ("xlsx", "xls") or result.get("kind") in _open.SHEETS
    whole_text = set()  # 文字の中に空欄のあるマス（inline）：読み戻した値は、マスの字の全体
    inline_before = result.pop("_inline", None) or set()
    for key, v in written.items():
        f = by_id.get(key, {})
        lab = labels.get(key) or f.get("label", "")
        if f.get("label") and f.get("context"):
            lab = f["label"]
        d = {"id": key, "label": lab, "value": f.get("value", v) if f else v}
        for k_ in ("sheet", "page"):
            if k_ in f:
                d[k_] = f[k_]
        if "sheet" not in d and "page" not in d and "!" in key and not key.startswith("@") and sheet_kind:
            # 欄の一覧にない、番地で書いたセル（「計画 (表)!BB3」）：番地のシートの所に入れる（ほかのシートの見出しの下にしない）
            d["sheet"] = key.rsplit("!", 1)[0].strip("'")
        if (f.get("kind") == "inline" or key in inline_before) and isinstance(d["value"], str) and "!" in key:
            whole_text.add(key)
        told.append(d)
    written_ids = {d["id"] for d in told} | set((result.get("anchored") or {}).values())
    said = said or {}
    # as を述べた鍵 → 伝える行（@ の書き込みは書いた所の id、セルの字の中の所「シート!B3@5」はそのセルの行）
    renamed_, anchored_ = result.get("renamed_ids") or {}, result.get("anchored") or {}
    rows = {d["id"]: d for d in told}
    for key, meaning in said.items():
        i = anchored_.get(renamed_.get(key, key), renamed_.get(key, key))
        d = rows.get(i) or rows.get(str(i).rsplit("@", 1)[0])
        if d is not None and meaning not in (d.get("as") or "").split("／"):
            d["as"] = f"{d['as']}／{meaning}" if d.get("as") else meaning
    rects = rects or {}
    written_rects = [rects[i] for i in written_ids if i in rects]

    def covered(f: dict) -> bool:
        """書いた所と同じ所の欄（同じ id・同じ所の別の読み方・ページの上で大きく重なる欄）：まだ空いている欄に数えない。"""
        if f["id"] in written_ids or f.get("alternative_to") in written_ids:
            return True
        r = rects.get(f["id"])
        return r is not None and any(_overlaps(r, w) for w in written_rects)

    empty = [f for f in after_list if f.get("value") in (None, "") and not f.get("readonly") and not f.get("agent")
             and f.get("kind") not in ("checkbox", "radio", "button", "picture") and not covered(f)]
    notes = list(result.get("warnings") or [])
    hand = result.get("hand_only")
    if hand:
        # 手で行う所（署名・押印など。段階の書く D の理由）：とどくは書いていない。本人が紙の上で行うことを、必ず伝える
        notes.insert(0, "手で行う所（とどくは書いていません。本人が紙の上で行います）："
                     + "／".join(f"{h['title']}" + (f"（{h['detail']}）" if h.get("detail") else "")
                                + (f"［{'・'.join(h['where'])}］" if h.get("where") else "") for h in hand))
    if not result.get("proof_ok", True):
        notes.insert(0, "書き込みの証明で、想定外の変更が見つかりました：" + "／".join(map(str, result.get("unexpected") or []))[:300])
    if not result.get("read_back_ok", True):
        notes.insert(0, "書いた後のファイルを読み直すと、書いた値と違う欄があります")

    def place(d: dict) -> str:
        where = d.get("sheet") or (f"{d['page']}ページ" if d.get("page") else "")
        return f"{where}・" if where else ""

    def bare_id(d: dict) -> str:
        i = d["id"]
        return i.split("!", 1)[1] if d.get("sheet") and i.startswith(d["sheet"] + "!") else i

    def source(d: dict) -> dict:
        return by_id.get(d["id"]) or d

    def own_name(d: dict) -> bool:
        """見出しが、位置から選んだ近くの字でなく、書類が欄に付けた名前か：データの列の名前・キー、PDF の欄の説明（/TU）や欄の名前、
        部品の Caption、Word などのコントロールの説明・タグ（印刷された字でないと窓が示したもの）。"""
        f, lab = source(d), d.get("label") or ""
        if not lab or "［" in lab:
            return False
        if result.get("kind") in _DATA_KINDS or f.get("tooltip") or f.get("label_method") == "部品の見出し（Caption）":
            return True
        if result.get("kind") == "pdf":
            return lab == d["id"].rsplit(".", 1)[-1]
        return not sheet_kind and d["id"] in labels and d["id"] not in printed and lab != d["id"]

    def near_text(d: dict) -> str:
        """近くの字の言い方（事実。欄の意味とはしない）。ない所は「近くの字なし」。窓の印は［ ］、書く所は「k つ目の空き」。"""
        lab = d.get("label") or ""
        if source(d).get("nearest_text"):
            return lab.split("（", 1)[1].rsplit("）", 1)[0] if "（" in lab else lab  # far_label：「左に 3 列あけて『氏名』」
        if not lab or (lab == d["id"] and d["id"] not in printed):
            return "近くの字なし"
        slots = list(re.finditer(r"［[^］]*］", lab))
        k = next((n for n, m in enumerate(slots, 1) if "▶" in m.group()), 0)
        base = re.sub(r"［[^］]*］", "［ ］", lab).strip("：:　 ") if len(slots) > 1 else re.sub(r"［[^］]*］", "", lab).strip("-－ー：:　 ")
        base = base if len(base) <= 40 else base[:40] + "…"  # 字のすべては窓に（ここは対話で読み上げる短い形）
        if not base:
            return "近くの字なし"
        return f"近くの字「{base}」" + (f"の {k} つ目の空き" if k and len(slots) > 1 else "")

    def name(d: dict) -> str:
        """読み上げても分かる所の名前：書類が欄に付けた名前か、id と近くの字。"""
        return (d.get("label") or "") if own_name(d) else f"{bare_id(d)}（{near_text(d)}）"

    def said_as(d: dict, v: str) -> str:
        """述べた意味があるとき：「意味：値（id／近くの字）」。"""
        tail = f"欄の名前「{d['label']}」" if own_name(d) else near_text(d)
        return f"{d['as']}：{v}（{bare_id(d)}／{tail}）"

    def shown(d: dict) -> str:
        v = "☑" if d["value"] is True else "☐" if d["value"] is False else _say_value(d["value"])
        if d.get("as"):
            return said_as(d, v.strip() if d["id"] in whole_text else v)
        if d["id"] in whole_text:
            return f"{bare_id(d)}：{v.strip()}"  # マスの字の全体（書いた後の字のとおり。空欄に差し込んだ形にしない）
        lab = d["label"] or ""
        m = re.search(r"［▶[^］]*］", lab) or re.search(r"［[ 　]*］", lab)
        if lab and m:
            rest = re.sub(r"［[^］]*］", "［ ］", lab[m.end():])
            head = re.sub(r"［[^］]*］", "［ ］", lab[:m.start()])
            return f"{bare_id(d)}：{head}［{v}］{rest}"  # 「令和［8］年［ ］月」：行の字のとおりに、空欄に値を入れた形で
        if own_name(d):
            return f"{lab.rstrip('：:　 ')}：{v}"  # 見出しの終わりの「：」は重ねない（「人数：12」）
        return f"{name(d)}：{v}"

    # 記入した内容の一覧：書いた所を、書類の中の順（シート・ページごと）に、全部。値は書いた後のファイルを読み直したもの
    order = {f["id"]: i for i, f in enumerate(after_list)}
    sheets: dict = {}
    for f in after_list:
        if f.get("sheet") is not None:
            sheets.setdefault(f["sheet"], len(sheets))
    # シートごとに（シートの順に）、その中は書類の中の順。欄の一覧にないセルは、そのシートの終わりに
    told.sort(key=lambda d: (sheets.get(d.get("sheet"), len(sheets)) if d.get("sheet") is not None else -1,
                             order.get(d["id"], len(order))))
    def joined(run: list[dict]) -> str:
        """1つの項目が分かれた欄（〒［▶］-［　］と〒［　］-［▶］）は、1行に：〒［123］-［4567］。"""
        lab = run[0]["label"]
        vals: dict[int, str] = {}
        for d in run:
            k = next((n for n, m in enumerate(re.finditer(r"［[^］]*］", d["label"])) if "▶" in m.group()), None)
            vals[k] = "☑" if d["value"] is True else "☐" if d["value"] is False else _say_value(d["value"])
        it = iter(range(10 ** 6))
        line = re.sub(r"［[^］]*］", lambda m: f"［{vals.get(next(it), ' ')}］", lab)
        ids = "・".join(bare_id(d) for d in run)
        return f"{run[0]['as']}：{line}（{ids}）" if run[0].get("as") else f"{ids}：{line}"

    def stem(d: dict) -> str | None:
        lab = d.get("label") or ""
        return (place(d) + re.sub(r"［[^］]*］", "［］", lab) + "\0" + (d.get("as") or "")) if "［▶" in lab and lab.count("［") > 1 else None

    lines = [f"記入した内容（{len(told)} か所。書いた後のファイルを読み直した値）："] if told or not (result.get("placed") or result.get("circled")) else []
    group = None
    i = 0
    while i < len(told):
        d = told[i]
        g = place(d).rstrip("・")
        if g and g != group:
            lines.append(f"【{g}】")
            group = g
        j = i + 1
        while stem(d) and j < len(told) and stem(told[j]) == stem(d):
            j += 1
        lines.append(f"- {joined(told[i:j]) if j - i > 1 else shown(d)}")
        i = j
    for pl in result.get("placed") or []:
        # 置いた画像（本人の印影・署名）：何を、どこに
        lines.append(f"{said[pl['key']]}：{pl['say']}" if pl.get("key") in said else pl["say"])
    for c in result.get("circled") or []:  # 囲んだ所は、いつも伝える（どの字を囲んだかは、上の一覧の値からは分からない）
        where = c.get("sheet") or (f"{c['page']}ページ" if c.get("page") else "")
        lab = c.get("label") or ""
        verb = "□で囲んだ（○ではなく、字の囲み線）" if c.get("style") == "border" else "○で囲んだ"
        meaning = said.get(c.get("key"))
        if meaning:
            c["as"] = meaning
        head = f"{meaning}：" if meaning else (f"近くの字「{lab}」：" if lab and lab != c["text"] else "")
        lines.append(f"- {verb}：{where + '・' if where else ''}{head}「{c['text']}」"
                     + (f"（前後の字：「{c['around']}」）" if c.get("around") else "")
                     + (f"（{c['uncertain']}）" if c.get("uncertain") else "")
                     + (f"（{c['inside_word']}）" if c.get("inside_word") else ""))
    if empty:
        # 次に書くとよい欄：書いた所の続き（同じシート・ページで、書いた所より後ろの空欄）から順に
        order = {f["id"]: i for i, f in enumerate(after_list)}
        last = max((order[d["id"]] for d in told if d["id"] in order), default=-1)
        where_last = next((after_list[last].get("sheet") or after_list[last].get("page") for _ in [0] if last >= 0), None)
        nxt = [f for f in empty if order.get(f["id"], -1) > last and (f.get("sheet") or f.get("page")) == where_last] or empty
        counts: dict = {}
        for f in empty:
            w = f.get("sheet") or (f"{f['page']}ページ" if f.get("page") else "")
            counts[w] = counts.get(w, 0) + 1
        by = "、".join(f"{w} {n} か所" for w, n in list(counts.items())[:6] if w) + ("ほか" if len(counts) > 6 else "")
        lines.append(f"まだ空いている欄：{len(empty)} か所" + (f"（{by}）" if by and len(counts) > 1 else ""))
        same = len({place(f) for f in nxt[:10]}) == 1
        head = f"続きの空欄（{place(nxt[0]).rstrip('・')}）：" if same and place(nxt[0]) else "続きの空欄："
        lines.append(head + "、".join(name(f) if same else f"{place(f)}{name(f)}" for f in nxt[:10]) + ("ほか" if len(nxt) > 10 else ""))
    # 選んでいない選択肢（チェックボックス・ラジオボタン）と、残っている「？」の所：数と、どこにあるか（一つずつは並べない）
    groups: dict = {}
    for f in after_list:
        if f.get("kind") not in ("checkbox", "radio") or f.get("readonly") or f.get("filled"):
            continue
        key = (place(f), f.get("group") or f.get("context") or f["id"])  # 同じ組（ラジオの組・同じ行の選択肢）
        on = f.get("value") not in (None, "", False) or f["id"] in written_ids or f.get("alternative_to") in written_ids
        groups[key] = groups.get(key, False) or on
    unchosen = [k_ for k_, on in groups.items() if not on]
    agent_left = [f for f in after_list if f.get("agent") and f.get("value") in (None, "") and not f.get("readonly")
                  and f.get("kind") not in ("checkbox", "radio") and not covered(f)]

    def by_where(keys: list[str]) -> str:
        cnt: dict = {}
        for w in keys:
            cnt[w] = cnt.get(w, 0) + 1
        named = [(w, n) for w, n in cnt.items() if w]
        return "（" + "、".join(f"{w} {n}" for w, n in named[:6]) + ("ほか" if len(named) > 6 else "") + "）" if len(named) > 1 else ""

    if unchosen:
        lines.append(f"選んでいない選択肢（チェックボックス・ラジオボタン）：{len(unchosen)} 組"
                     + by_where([w.rstrip("・") for w, _g in unchosen]))
    if agent_left:
        lines.append(f"規則では決められない「？」の所：{len(agent_left)} か所"
                     + by_where([place(f).rstrip("・") for f in agent_left]) + "（書く所かは、窓を読んで判断する）")
    if result.get("recalculated_cells"):
        from .xlsx.readback import say as _recalc_say

        lines.append(_recalc_say(result["recalculated_cells"]))  # 書いた値から計算し直して変わったセル（金額・合計など）
    if notes:
        lines.append("本人に伝えること：" + "／".join(notes[:10]))
    # as のない書き込み：本人が確かめる書類なら、書いた値の意味を as で述べて書き直せる（対話が本人の確かめる手段なので）
    without_as = [d["id"] for d in told if not d.get("as")] + [c.get("id") or c.get("key") for c in result.get("circled") or [] if not c.get("as")]
    return {"written": told, "written_without_as": without_as, "still_empty": [{"id": f["id"], "label": f.get("label", ""), **({"sheet": f["sheet"]} if "sheet" in f else {}),
                                              **({"page": f["page"]} if "page" in f else {})} for f in empty[:50]],
            "still_empty_count": len(empty), "unchosen_choice_groups": len(unchosen), "agent_places_left": len(agent_left),
            "say_to_user": "\n".join(lines)}


@_budgeted
def run_macro(path: str, macro: str, out_path: str, answers: list | None = None, files: dict | None = None,
              password: str | None = None, overwrite: bool = False) -> dict:
    """マクロ（ボタンの処理など）を動かし、マクロが変えたセルを証明つきで out_path に書く。
    answers：MsgBox・InputBox の答え（順に）。files：マクロが読むファイル（マクロの中のパス → 手元のファイル）。
    人の答えやファイルが要るときは、勝手に決めずに、その問いを返す。"""
    from .book import open_book
    from .formula.parser import addr
    from .vbarun import NeedsInput, Unsupported
    from .vbarun import run_macro as run
    from .vbarun.parse import VbaSyntaxError
    from .xlsx.edit import EditError
    from .xlsx.edit import fill as xfill
    from .xlsx.proof import prove

    src = Path(path).expanduser().resolve()
    dst = Path(out_path).expanduser().resolve()
    _not_source(src, dst)
    blob, name, k, lock = _read(path, password)
    if k not in ("xlsx", "xls"):
        raise ApiError("マクロを動かせるのは、いまは Excel の書類です")
    wb = open_book(blob, name)
    inputs = {key: Path(v).expanduser().read_bytes() for key, v in (files or {}).items()}
    try:
        r = run(wb, macro, answers, inputs)
    except NeedsInput as e:
        what = {"msgbox": "はい・いいえなどの答え", "inputbox": "入力する値", "file": "ファイル"}.get(e.kind, e.kind)
        return {"done": False, "needs": e.kind, "prompt": e.prompt, "line": e.line,
                "note": f"マクロが{what}を求めています。本人に確かめて、answers（または files）に入れてもう一度呼んでください"}
    except (Unsupported, VbaSyntaxError) as e:
        return {"done": False, "unsupported": str(e), "note": "この命令は、規則では再現できません。ソースを読んで判断してください"}
    out: dict = {"done": True, "macro": macro, "messages": r.messages, "printed": r.printed, "error": r.error,
                 "volatile": r.volatile, "changed": {f"{s_}!{addr(rr, cc)}": v for (s_, rr, cc), v in r.changed.items()}}
    if r.notes:
        out["notes"] = list(r.notes)
    if r.files:
        fdir = dst.parent / f"{dst.stem}_マクロのファイル"
        fdir.mkdir(parents=True, exist_ok=True)
        saved = []
        for fname, data in r.files.items():
            target = fdir / Path(fname.replace("\\", "/")).name
            target.write_bytes(data)
            saved.append(str(target))
        out["files"] = saved
    if not r.changed:
        out["note"] = "マクロは、セルの値を変えませんでした（ファイルは書き出していません）"
        return out
    values = {f"{s_}!{addr(rr, cc)}": v for (s_, rr, cc), v in r.changed.items()}
    try:
        res = xfill(open_book(blob, name), values, allow_formula_overwrite=True, force=True, by_macro=set(r.changed),
                    run_macros=False)
    except EditError as e:
        raise ApiError(str(e)) from None
    pr = prove(blob, res.blob, res.allowed_inputs, {(p.sheet, p.address) for p in res.patches if p.kind == "control"},
               {(p.sheet, p.address) for p in res.patches if p.kind in ("macro", "input")})
    out_blob = res.blob
    if lock is not None:
        from .unlock import relock, unlock

        out_blob = relock(lock, res.blob)
        if unlock(out_blob, password)[0] != res.blob:
            raise ApiError("保護し直したファイルを同じパスワードで解くと、書いた中身と一致しません（書き出していません）")
        out["password_protected"] = f"元と同じパスワードで保護したまま書き出しました（{lock.label}）"
    out.update({"out_path": str(dst), "proof_ok": pr.ok, "proof": pr.headline(), "unexpected": pr.unexpected[:20],
                "recalculated": {f"{p.sheet}!{p.address}": p.after for p in res.patches if p.kind == "formula"},
                "warnings": res.warnings})
    tmp = _temp_beside(dst, out_blob)
    try:
        _publish(tmp, dst, src, overwrite, out, ok=pr.ok)
    finally:
        tmp.unlink(missing_ok=True)
    return out


def _num_equal(a, b) -> bool:
    try:
        return abs(float(a) - float(b)) < 1e-9
    except (TypeError, ValueError):
        return False


@_budgeted
def audit(path: str, password: str | None = None, encoding: str | None = None) -> dict:
    import inspect

    from .unlock import pdf_password

    blob, name, k, lock = _read(path, password)
    encoding = _encoding(blob, k, encoding)
    if k in _arc.KINDS:
        return _audit_zip(path, blob, password, k, name)
    if k == "pdf":
        from .pdf.audit import run
    elif k == "docx":
        from .docx.audit import run
    elif k == "doc":
        from .doc.audit import run
    elif k == "jtd":
        from .jtd.audit import run
    elif k in _open.KINDS:  # 追加の形式
        run = _open.audit_runner(k, encoding)
    else:
        from .xlsx.audit import run
    pw = pdf_password(lock)
    au = run(blob, name, password=pw) if pw and "password" in inspect.signature(run).parameters else run(blob, name)
    d = au.to_dict()
    from . import pkgsig

    sig = pkgsig.find(blob)
    if sig is not None:
        # パッケージの電子署名：確かめ（audit）の中の書き込みの試しは、署名を無効にする（書き出したファイルは署名し直しが要る）
        d["digital_signature"] = sig.to_dict()
        d.setdefault("checks", []).append({"name": "電子署名", "ok": True, "agent": False,
                                           "detail": sig.message + "。署名そのものの正しさ（証明書・ハッシュ）は、とどくは確かめていません",
                                           "items": list(sig.parts)})
    d["grade"] = grade(path, password=password, encoding=encoding)
    if lock is not None:
        d["password_protected"] = f"パスワードで保護された書類です（{lock.label}）。解いた中身を確かめました"
    return d


@_budgeted
def read_table(path: str, header: str, sheet: str | None = None, header_rows: int | None = None,
               out_path: str | None = None, password: str | None = None) -> dict:
    """Excel の表を、見出しの列で行ごとに取り出す（結合セルは画面どおりに範囲へ広げ、値ごとに元のセルを添える）。
    sheet を省くと、header のあるすべてのシートから。out_path（.csv か .json）を指すと、すべての行を書き出す。"""
    import csv
    import io
    import json

    from .book import open_book
    from .xlsx.tables import TableError
    from .xlsx.tables import read as tread

    blob, name, k, lock = _read(path, password)
    if k not in _BOOKS:
        raise ApiError("表の取り出しは Excel（.xlsx・.xlsm・.xls）です（ほかの形式は窓で読めます）")
    try:
        t = tread(open_book(blob, name), header, sheet, header_rows)
    except TableError as e:
        raise ApiError(str(e)) from None
    out = {"columns": t["columns"], "rows": len(t["rows"]), "sheets": t["sheets"][:50], "sheet_count": len(t["sheets"]),
           "first_rows": t["rows"][:20], "check": t["check"]}
    if out_path:
        dst = Path(out_path).expanduser().resolve()
        src = Path(_member(path)[0] if _member(path) else path).expanduser().resolve()
        if dst == src:
            raise ApiError("元のファイルは上書きしません。別の名前（out_path）を指定してください")
        dst.parent.mkdir(parents=True, exist_ok=True)
        if dst.suffix.lower() == ".json":
            dst.write_text(json.dumps(t, ensure_ascii=False, indent=1), encoding="utf-8")
        else:
            buf = io.StringIO()
            w = csv.writer(buf, lineterminator="\n")
            head = ["シート", "行"] + t["columns"]
            w.writerow(head)
            for r in t["rows"]:
                w.writerow([r["_sheet"], r["_row"]] + [r.get(c, "") for c in t["columns"]])
            dst.write_text(buf.getvalue(), encoding="utf-8")
        out["out_path"] = str(dst)
    else:
        out["all_rows"] = t["rows"]
    return out


@_budgeted
def calc_model(path: str, cell: str | None = None, password: str | None = None) -> dict:
    """Excel の様式の「データ」（入力欄）と「処理」（入力につながる数式）。cell を指すと、その計算を入力欄までさかのぼる。
    式は、番地を見出しに置き換えて示す（in_words）。"""
    from .book import open_book
    from .xlsx import model as xm

    blob, name, k, lock = _read(path, password)
    if k not in _BOOKS:
        raise ApiError("計算の取り出しは Excel の様式（.xlsx・.xlsm・.xls）です")
    _k, doc, _ = window(path, password)
    m = xm.model(open_book(blob, name), doc)
    if cell:
        try:
            return {"explain": xm.explain(m, cell)}
        except ValueError as e:
            raise ApiError(str(e)) from None
    return {"inputs": len(m["inputs"]), "steps": len(m["steps"]), "results": len(m["results"]),
            "dynamic": m["dynamic"], "result_groups": xm.summary(m)}


@_budgeted
def calculate(path: str, values: dict, password: str | None = None) -> dict:
    """値の組から、ファイルを書かずに結果を計算する（試算）。書き込みと同じ受け取り方・入力規則・計算エンジン。"""
    from .book import open_book
    from .xlsx import model as xm
    from .xlsx.edit import EditError

    blob, name, k, lock = _read(path, password)
    if k not in _BOOKS:
        raise ApiError("試算は Excel の様式（.xlsx・.xlsm・.xls）です")
    _k, doc, _ = window(path, password)
    values = _split_values(values, k)[0]  # as（述べた意味）は書くときのもの。試算にも同じ値の形を渡せるように
    try:
        r = xm.calculate(open_book(blob, name), doc, values)
    except EditError as e:
        raise ApiError(str(e)) from None
    lines = [f"試算（ファイルは書いていません）：入れた値 {len(r['inputs'])} か所で、計算が変わる所 {len(r['changed'])} か所。"]
    for c in r["changed"][:40]:
        lines.append(f"- {c['label'] or c['cell']}：{c['before'] or '（空）'} → {c['after'] or '（空）'}")
    if r["not_computed"]:
        lines.append(f"計算し直せなかった所 {len(r['not_computed'])} か所（Excel で開いたときに計算されます）："
                     + "、".join(x["cell"] for x in r["not_computed"][:10]))
    r["say_to_user"] = "\n".join(lines)
    return r


@_budgeted
def view_page(path: str, page: int, region: list | None = None, marks: list | None = None,
              password: str | None = None, dpi: float | None = None) -> dict:
    """PDF のページの画像（スキャンした紙のページを、エージェントが見て読み、書く所を mm の枠で決めるため）。
    todoku が書き込んだ欄（p ページ.w 番号など）は、青い枠で描き、id と値を添える。欄・注釈の見た目（/AP）の字は、見た目の
    とおりの位置・大きさ・向きで描く（pdf/pageimage.py）。dpi：解像度（省くとページ全体は 100 dpi。画素の数の上限は pageimage のとおり）。"""
    from .pdf.pageimage import view_page as pview

    blob, name, k, lock = _read(path, password)
    if k == "image":
        from .imaging.standalone import view_frame

        try:
            if dpi is not None and (isinstance(dpi, bool) or not isinstance(dpi, (int, float)) or not 10 <= dpi <= 1200):
                raise ApiError("dpi は 10〜1200 の数で指定してください")
            return view_frame(blob, int(page), region, marks, dpi)
        except ValueError as e:
            raise ApiError(str(e)) from None
        except (struct.error, IndexError) as e:
            raise ApiError(broken_message(path, e)) from None
    if k != "pdf":
        raise ApiError("view_page は PDF と画像のファイルのページの画像です（ほかの形式は窓で読み、書類の中の画像は view_image で見ます）")
    doc = _pdf(blob, name, lock)
    try:
        if dpi is not None and (isinstance(dpi, bool) or not isinstance(dpi, (int, float)) or not 10 <= dpi <= 1200):
            raise ApiError("dpi は 10〜1200 の数で指定してください（省くとページ全体は 100 dpi）")
        from .limits import LimitExceeded

        problem = ""
        try:
            written = _written_boxes(path, password, int(page))
        except LimitExceeded:
            raise
        except Exception as e:  # noqa: BLE001  窓を作れない書類でも、ページの画像は出す（青い枠だけ描けないと伝える）
            written, problem = None, f"値の入った欄の青い枠は描けません（窓を作れません：{type(e).__name__}）"
        v = pview(doc, int(page), region, marks, written=written, dpi=dpi)
        if problem:
            v["text"] += " " + problem + "。"
            v["not_drawn"].append(problem)
        return v
    except ValueError as e:
        raise ApiError(str(e)) from None


def view_image(path: str, image: str, password: str | None = None, region: list | None = None) -> dict:
    """書類の中の画像を渡す（窓の「書類の中の画像」の id か、部品の名前で指す）。PNG・JPEG はそのまま、ほかは描いて PNG で。
    結果：{id, mime, data_base64, width, height, not_drawn, text}。"""
    from .imaging.docimages import ImageNotFound, get_image, pdf_get

    blob, name, k, lock = _read(path, password)
    try:
        if k == "pdf":
            return pdf_get(_pdf(blob, name, lock), image, region)
        if k == "image":
            raise ApiError("画像のファイルは view_page で見ます")
        _k, doc, _parts = window(path, password)
        return get_image(k, blob, doc.get("images") or [], image, region)
    except ImageNotFound as e:
        raise ApiError(str(e)) from None
    except ApiError:
        raise
    except (ValueError, struct.error, IndexError, KeyError) as e:
        raise ApiError(f"画像を開けません：{e}") from None


def to_pdf(path: str, out_path: str, overwrite: bool = False) -> dict:
    """画像のファイルを PDF に包む（元のファイルは変えない）。できた PDF に fill_document で枠（mm）を指定して書く。"""
    from .imaging.standalone import to_pdf as wrap

    blob, _name, k, _lock = _read(path)
    if k != "image":
        raise ApiError("to_pdf は画像のファイルを PDF に包む道具です（この書類はそのまま読み書きできます）")
    src, dst = Path(path).expanduser(), Path(out_path).expanduser()
    _not_source(src, dst)
    if dst.exists() and not overwrite:
        raise ApiError(f"書き出し先のファイルがすでにあります：{dst.name}（overwrite で上書きできます）")
    pdf, pages, notes = wrap(blob)
    dst.write_bytes(pdf)
    return {"out_path": str(dst), "pages": pages, "notes": notes,
            "say_to_user": f"画像を PDF（{len(pages)} ページ）に包みました。書くのはこの PDF です（元の画像は変えていません）。" + " ".join(notes)}


def _pdf_window_doc(path: str, password: str | None) -> dict:
    """PDF の窓の JSON だけ。窓の Markdown を作らない（描くのに要らないので、Markdown の組み立ての失敗で画像まで出せなくならない）。"""
    import hashlib
    import pickle

    from .pdf.window import build

    blob, name, _k, lock = _read(path, password)
    key = (hashlib.sha256(blob).hexdigest(), name, None) if lock is None else None
    if key is not None and key in _WINDOWS:
        return pickle.loads(_WINDOWS[key][1])[1]
    return build(_pdf(blob, name, lock))


def _written_boxes(path: str, password: str | None, page: int) -> list[dict]:
    """ページの、値の入った入力欄（フォームの欄・todoku が書いた欄）の枠（mm）と値。"""
    from .pdf.layout import MM_PT

    out = []
    doc = _pdf_window_doc(path, password)
    for pg in doc.get("pages", []):
        if pg["number"] != page:
            continue
        for f in pg["fields"]:
            v = f.get("value")
            r = f.get("rect")
            if v in (None, "", False) or not r:
                continue
            mm = MM_PT / (pg.get("user_unit") or 1.0)  # 1 mm のページの座標の長さ（UserUnit も反映）
            out.append({"box": [r[0] / mm, r[1] / mm, r[2] / mm, r[3] / mm], "id": f["id"], "value": v})
    return out


def _audit_zip(path: str, blob: bytes, password: str | None, k: str = "zip", name: str = "") -> dict:
    """ZIP：中の、扱える形式のファイルを1つずつ検査する（検査の名前の頭に、中の名前を付ける）。"""
    from . import ziparc

    checks = []
    if k in _arc.MAIL:
        _kk, doc, _parts = window(path, password)
        from .readonlydoc import chars

        missing = chars(_arc.mail_body(blob, k)) - chars("".join(doc["body_lines"]))
        checks.append({"name": "読み落とし（本文）", "ok": not missing, "agent": False, "items": [str(x) for x in list(missing)[:20]],
                       "detail": "差出人・宛先・件名・本文の字が、すべて窓に出ていることを確かめました" if not missing else "本文の字のうち窓に出ていない字があります"})
    for m in _arc.files(blob, k, name):
        try:
            d = audit(path + ziparc.SEP + m.name, password)
        except ApiError as e:
            checks.append({"name": f"{m.name}：開けない", "ok": False, "agent": False, "detail": str(e), "items": []})
            continue
        checks += [{**c, "name": f"{m.name}：{c['name']}"} for c in d["checks"]]
    return {"file": Path(path).name, "ok": all(c["ok"] for c in checks), "needs_agent": any(c["agent"] for c in checks),
            "checks": checks, "grade": grade(path, password=password)}


@_budgeted
def grade(path: str, findings: list[dict] | None = None, answered: list[str] | None = None,
          password: str | None = None, encoding: str | None = None) -> dict:
    """段階（S〜F）：誰が扱えば正確に読んで入れられるか。規則で決めた段階と、エージェントが確かめること。
    findings（エージェントが窓を読んで見つけたこと：{grade, title, detail, where}）を渡すと、段階に合わせる。
    パスワードで保護された書類は、本人に聞いたパスワードで解いた中身で決める（パスワードがなければ PasswordNeeded）。"""
    from .cfb import CfbError
    from .grade import combine, grade_blob, unsupported_grade
    from .unlock import pdf_password

    try:
        blob, name, k, lock = _read(path, password)
    except Unsupported as e:
        g = unsupported_grade(e.kind, detail=str(e))  # まだ読めない形式：段階 F と、その理由
        return (combine(g, findings or [], answered) if findings or answered else g).to_dict()
    except CfbError as e:
        # 形式を見分ける前に、複合ファイルの作りが壊れている：開けない（F）
        from .grade import _locked

        g = _locked("開けない", str(e), "読めないので、書けません")
        return (combine(g, findings or [], answered) if findings or answered else g).to_dict()
    encoding = _encoding(blob, k, encoding)
    g = grade_blob(blob, name, password=pdf_password(lock), encoding=encoding)
    if findings or answered:
        g = combine(g, findings or [], answered)
    _remember_hand_only(path, g, findings, answered)
    return g.to_dict()


# 手で行う所（署名・押印・拇印・公証など。書く段階 D の理由）：段階を決めたときに見つかったものを、ファイルの中身ごとに控え、
# そのファイルに書いた結果の「本人に伝えること」に入れる（白紙のまま残る署名の欄を、本人に伝え落とさないように）
_HAND_ONLY: "dict[str, list[dict]]" = {}


def _remember_hand_only(path: str, g, findings, answered) -> None:
    hand = [{"title": r.title, "detail": r.detail, "where": list(r.where), "by": r.by}
            for r in g.write_reasons if r.grade == "D" and r.by == "agent"]
    checked = bool(hand) or "cannot_do" in (answered or []) or any((f or {}).get("check") == "cannot_do" for f in findings or [])
    if not checked:
        return
    p = Path(path).expanduser()
    if not p.is_file():
        return
    try:
        _HAND_ONLY[hashlib.sha256(_read_file(p)).hexdigest()] = hand
    except Exception:  # noqa: BLE001  控えられないだけ（書き込みの結果に「まだ確かめていない」と出る）
        return
    while len(_HAND_ONLY) > 64:
        _HAND_ONLY.pop(next(iter(_HAND_ONLY)))


@_budgeted
def summary(path: str, password: str | None = None, encoding: str | None = None) -> dict:
    k, doc, parts = window(path, password, encoding)
    fl = fields(path, password, encoding)
    d = {"file": Path(path).name, "kind": k, "kind_ja": KIND_JA[k], "fields": len(fl),
         "writable_fields": sum(1 for f in fl if not f.get("readonly")), "window_parts": len(parts),
         "window_chars": sum(len(t) for _, t in parts)}
    if k == "pdf":
        d["pages"] = len(doc["pages"])
        d["form_fields"] = doc["form"]["acroform_fields"]
    elif k in ("xlsx", "xls"):
        d["sheets"] = [s["name"] for s in doc["sheets"]]
        m = doc.get("macros")
        d["macros"] = bool(m)
        if isinstance(m, dict) and m.get("readable"):
            d["macro_procedures"] = sum(len(x["procedures"]) for x in m["modules"])
            d["macro_runs_on_change"] = any(e["on_change"] for e in m["events"])
    elif k in _arc.KINDS:
        d["members"] = [{"name": m["name"], "kind": m["kind"], "size": m["size"]} for m in doc["members"]]
    else:
        d["protection"] = doc.get("protection")
    if k in _open.KINDS:  # 追加の形式
        d.pop("protection", None)
        d.update(_open.summary(k, doc))
    return d
