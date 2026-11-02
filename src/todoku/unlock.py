"""パスワードで保護された書類：本人が対話で伝えたパスワードで開き、書いた後も同じパスワードで保護したまま書き出す。

- Office の新しい形（.xlsx・.docx を暗号の入れ物に入れたもの：Agile・Standard）は officecrypto で解き、かけ直す
- 古い形（.xls・.doc の RC4・RC4 CryptoAPI・XOR）は bincrypto で解き、かけ直す
- PDF は、文書を開くときにパスワードを渡す（PDF の部品が中で解き、書き足す所も同じ鍵で暗号にする）

パスワードは、結果にも記録にも残さない（この入口の外へは、解いた中身とかけ直すための情報だけを渡す）。
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .book import kind_of


class NeedsPassword(ValueError):
    """開くのにパスワードが要る（本人に聞いて、もう一度呼ぶ）。"""


class WrongPassword(ValueError):
    """パスワードが違う（本人にもう一度確かめる）。"""


@dataclass
class Lock:
    method: str  # "ooxml"（Office の新しい形）/ "binary"（古い .xls・.doc）/ "pdf"
    key: object = None  # かけ直すための情報（パスワードそのものは持たない。PDF は文書を開くときに使う）
    outer_kind: str = ""  # 暗号の入れ物の形（"encrypted"・"xls"・"doc"・"pdf"）
    _password: str | None = field(default=None, repr=False)  # PDF を開き直すときだけ使う（結果には出さない）

    @property
    def label(self) -> str:
        return {"ooxml": "Office の暗号（パスワードで開く）", "binary": "古い形の暗号（パスワードで開く）",
                "pdf": "PDF の暗号（パスワードで開く）"}.get(self.method, self.method)


def needs_password(blob: bytes) -> bool:
    """開くのにパスワードが要る書類か（空のパスワードで開ける PDF は、要らない）。"""
    k = kind_of(blob)
    if k == "encrypted":
        return True
    if k in ("xls", "doc"):
        bincrypto = _bincrypto()
        if bincrypto is None:
            return False
        return bincrypto.is_encrypted_xls(blob) if k == "xls" else bincrypto.is_encrypted_doc(blob)
    if k == "pdf" and b"/Encrypt" in blob:
        from .pdf.doc import PdfDocument

        try:
            PdfDocument(blob, "")
        except Exception as e:  # noqa: BLE001
            return _is_pdf_password_error(e)
    return False


def unlock(blob: bytes, password: str | None) -> tuple[bytes, Lock | None]:
    """(読む中身, かけ直すための情報)。暗号がなければ、そのままと None。パスワードが要るのに無ければ NeedsPassword。"""
    k = kind_of(blob)
    if k == "encrypted":
        if not password:
            raise NeedsPassword("この書類はパスワードで保護されています（開くのにパスワードが要ります）")
        from . import officecrypto

        try:
            inner, key = officecrypto.decrypt_ooxml(blob, password)
        except officecrypto.WrongPassword:
            raise WrongPassword("パスワードが違います") from None
        except officecrypto.CryptoError as e:
            raise ValueError(str(e)) from None
        return inner, Lock("ooxml", key, k)
    if k in ("xls", "doc"):
        bincrypto = _bincrypto()
        if bincrypto is None or not (bincrypto.is_encrypted_xls(blob) if k == "xls" else bincrypto.is_encrypted_doc(blob)):
            return blob, None
        if not password:
            raise NeedsPassword("この書類はパスワードで保護されています（開くのにパスワードが要ります）")
        try:
            plain, key = bincrypto.decrypt_binary(blob, password)
        except bincrypto.WrongPassword:
            raise WrongPassword("パスワードが違います") from None
        except bincrypto.CryptoError as e:
            raise ValueError(str(e)) from None
        return plain, Lock("binary", key, k)
    if k == "pdf" and b"/Encrypt" in blob:  # 暗号の辞書がない PDF は、開き直して確かめるまでもない
        from .pdf.doc import PdfDocument

        try:
            PdfDocument(blob, "")
            return blob, None
        except Exception as e:  # noqa: BLE001
            if not _is_pdf_password_error(e):
                raise
        if not password:
            raise NeedsPassword("この PDF はパスワードで保護されています（開くのにパスワードが要ります）")
        try:
            PdfDocument(blob, "", password=password)
        except Exception as e:  # noqa: BLE001
            if _is_pdf_password_error(e):
                raise WrongPassword("パスワードが違います") from None
            raise
        return blob, Lock("pdf", None, k, password)
    return blob, None


def relock(lock: Lock | None, plain: bytes) -> bytes:
    """書いた中身を、元と同じパスワードで保護し直す（PDF は書き足す所が既に暗号なので、そのまま）。"""
    if lock is None or lock.method == "pdf":
        return plain
    if lock.method == "ooxml":
        from . import officecrypto

        return officecrypto.encrypt_ooxml(lock.key, plain)
    from . import bincrypto

    return bincrypto.encrypt_binary(lock.key, plain)


def pdf_password(lock: Lock | None) -> str | None:
    """PDF を開き直すときのパスワード（この入口と api の中だけで使う）。"""
    return lock._password if lock is not None and lock.method == "pdf" else None


def _bincrypto():
    """古い形（.xls・.doc）の暗号の部品（まだ無いときは None：暗号の確かめを飛ばす）。"""
    try:
        from . import bincrypto
    except ImportError:
        return None
    return bincrypto


def _is_pdf_password_error(e: Exception) -> bool:
    name = type(e).__name__
    return name in ("NeedsPassword", "WrongPassword", "PasswordRequired", "PdfPasswordError") or "パスワード" in str(e)
