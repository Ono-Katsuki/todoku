"""文字のファイルの文字コード：決まった順に試して決める（同じファイルなら、いつも同じ答え）。

書き込みはバイトの位置で行うので、字の位置（文字列の中の番号）とバイトの位置を結ぶ。Shift_JIS（CP932）には、
同じ字に2つのバイト列があるもの（NEC 選定 IBM 拡張文字と IBM 拡張文字など）があるので、ファイル全体を文字列にして
書き直すことはしない（書いた所だけを、そのファイルの文字コードで作ったバイト列に置き換える）。
"""

from __future__ import annotations

from dataclasses import dataclass

BOMS = [(b"\xef\xbb\xbf", "utf-8"), (b"\xff\xfe", "utf-16-le"), (b"\xfe\xff", "utf-16-be")]
# BOM がないときに試す順：UTF-8（ASCII だけのファイルもここで決まる）→ Shift_JIS（CP932）→ EUC-JP。どれでも読めなければ Latin-1
TRY = ("utf-8", "cp932", "euc_jp")
NAME_JA = {"utf-8": "UTF-8", "cp932": "Shift_JIS（CP932）", "euc_jp": "EUC-JP", "utf-16-le": "UTF-16（リトルエンディアン）",
           "utf-16-be": "UTF-16（ビッグエンディアン）", "latin-1": "Latin-1（文字コードを決められなかった）"}


class EncodingError(ValueError):
    pass


@dataclass
class Decoded:
    encoding: str
    bom: bytes
    text: str
    ascii_only: bool  # ASCII の字だけ（文字コードは UTF-8 とみなしたが、ファイルからは決まらない）

    @property
    def label(self) -> str:
        return NAME_JA.get(self.encoding, self.encoding) + ("・BOM あり" if self.bom else "・BOM なし")

    def byte_of(self, i: int) -> int:
        """文字列の i 番目の字の始まりの、ファイルの中のバイトの位置（字ごとのバイト数を足し上げた表を、1回だけ作る）。"""
        pre = getattr(self, "_prefix", None)
        if pre is None:
            sizes: dict[str, int] = {}
            pre = [0] * (len(self.text) + 1)
            at = 0
            for k, ch in enumerate(self.text):
                if ch < "\x80" and self.encoding not in ("utf-16-le", "utf-16-be"):
                    at += 1
                else:
                    n = sizes.get(ch)
                    if n is None:
                        n = sizes[ch] = len(ch.encode(self.encoding))
                    at += n
                pre[k + 1] = at
            self._prefix = pre
        return len(self.bom) + pre[i]

    def span(self, blob: bytes, a: int, b: int) -> tuple[int, int]:
        """字の範囲 [a, b) → バイトの範囲。読み直して同じ字になることを確かめる。"""
        x, y = self.byte_of(a), self.byte_of(b)
        if blob[x:y].decode(self.encoding) != self.text[a:b]:
            raise EncodingError("字の位置とバイトの位置を結べませんでした")
        return x, y

    def encode(self, s: str) -> bytes:
        try:
            return s.encode(self.encoding)
        except UnicodeEncodeError as e:
            bad = s[e.start:e.end]
            raise EncodingError(f"「{bad}」は、このファイルの文字コード（{NAME_JA.get(self.encoding, self.encoding)}）では書けない字です") from None


def decode(blob: bytes) -> Decoded:
    for bom, enc in BOMS:
        if blob.startswith(bom):
            return Decoded(enc, bom, blob[len(bom):].decode(enc, errors="strict"), False)
    for enc in TRY:
        try:
            text = blob.decode(enc)
        except UnicodeDecodeError:
            continue
        return Decoded(enc, b"", text, enc == "utf-8" and blob.isascii())
    return Decoded("latin-1", b"", blob.decode("latin-1"), False)
