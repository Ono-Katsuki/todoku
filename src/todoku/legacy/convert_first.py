"""とどくが自分では解かない形式（RAR・HEIC／HEIF・AVIF）：見分けて、別の道具で変換・展開する手順を返す（kind "rar"・"heic"・"avif"）。

とどくはここで何も実行しない（外の道具を呼ばない）。手順は目安で、どの道具が入っているかはエージェントが確かめる。
段階は読む E（別の道具で変換してから読む）・書く D。変換してできたファイルを、もう一度開けば、その形式として読み書きできる。
"""

from __future__ import annotations

import shlex
from pathlib import PurePath

from .. import readonlydoc as rd

KIND_JA = {"rar": "RAR の書庫（.rar。別の道具で展開してから読む）", "heic": "HEIC・HEIF の画像（別の道具で変換してから読む）", "avif": "AVIF の画像（別の道具で変換してから読む）"}
FORMAT = {"rar": "RAR", "heic": "HEIC（HEIF）", "avif": "AVIF"}
RAR_MESSAGE = "RAR の中身は別の道具で展開してください（展開したフォルダか ZIP なら読めます）"
WHY = {"rar": "RAR の圧縮の形には公開の仕様がなく、ライセンスの条件もあるので、とどくは RAR を解きません",
       "heic": "HEIC の画像の符号（HEVC）には特許の許諾が関わるので、とどくは HEIC を解きません",
       "avif": "AVIF の画像の符号（AV1）は特許の条件の確認が要るので、とどくは AVIF を解きません"}
AFTER = "変換・展開してできたファイルを、もう一度 open_document で開いてください"
WRITES = "書いたものは元の形式には戻りません（PNG や展開したファイルに対して書きます）。本人に伝えてください"
HINT = "どれも入っているとは限らない目安です。入っている道具を確かめて、上から順に試してください。変換・展開したことは本人に伝えてください"
WRITE = "この形式には書きません。" + WRITES


def kind_of_blob(blob: bytes) -> str | None:
    """先頭のバイトで見分ける（book.kind_of と同じ決め方）。"""
    if blob[:6] == b"Rar!\x1a\x07":
        return "rar"
    from ..imaging.formats import heif_brand

    brand = heif_brand(blob)
    return "avif" if brand == "AVIF" else "heic" if brand == "HEIC" else None


def _image_commands(src: str, out: str, avif: bool) -> list[dict]:
    q, o = shlex.quote(src), shlex.quote(out)
    rows = [{"os": "macOS", "command": f"sips -s format png {q} --out {o}", "note": "macOS に初めから入っています（AVIF は新しい macOS で）"},
            {"os": "macOS・Linux・Windows", "command": f"magick {q} {o}", "note": "ImageMagick"},
            {"os": "macOS・Linux・Windows", "command": f"ffmpeg -i {q} {o}", "note": "古い ffmpeg は、タイルに分かれた写真の 1 枚だけを書き出すことがあります。できた画像を view_page で見て確かめてください"}]
    if avif:
        rows += [{"os": "macOS・Linux", "command": f"avifdec {q} {o}", "note": "libavif"},
                 {"os": "macOS・Linux", "command": f"heif-convert {q} {o}", "note": "libheif"},
                 {"os": "macOS・Linux・Windows", "command": f"""python3 -c "import sys; from PIL import Image; Image.open(sys.argv[1]).save(sys.argv[2])" {q} {o}""", "note": "AVIF を読める版の Pillow が入っているとき"}]
    else:
        rows += [{"os": "macOS・Linux", "command": f"heif-convert {q} {o}", "note": "libheif"},
                 {"os": "macOS・Linux・Windows", "command": f"""python3 -c "import sys, pillow_heif; from PIL import Image; pillow_heif.register_heif_opener(); Image.open(sys.argv[1]).save(sys.argv[2])" {q} {o}""", "note": "pillow-heif が入っているとき"}]
    rows.append({"os": "Windows", "command": "（本人の操作）「フォト」で開いて、PNG か JPEG で保存する", "note": "Microsoft Store の画像拡張機能が要ることがあります"})
    return rows


def _rar_commands(src: str, folder: str) -> list[dict]:
    q, d = shlex.quote(src), shlex.quote(folder)
    return [{"os": "macOS・Linux", "command": f"mkdir -p {d} && bsdtar -xf {q} -C {d}", "note": "macOS に初めから入っています"},
            {"os": "Windows", "command": f"mkdir {d}; tar -xf {q} -C {d}", "note": "新しい Windows の tar は RAR も読めます"},
            {"os": "macOS・Linux", "command": f"unar -o {d} {q}", "note": "The Unarchiver のコマンド"},
            {"os": "macOS・Linux・Windows", "command": f"unrar x {q} {d}/", "note": "UnRAR"},
            {"os": "macOS・Linux・Windows", "command": f"7z x {q} -o{d}", "note": "7-Zip（7zz の版によっては RAR を解けません）"},
            {"os": "すべて", "command": "（パスワードがあるとき）本人にパスワードを聞いて、その道具の指定で渡す", "note": ""}]


def needs_tool(kind: str, name: str = "") -> dict:
    """format・why・next_step（commands・after・writes・hint）。"""
    base = PurePath(name.replace("\\", "/")).name or ("a.rar" if kind == "rar" else f"a.{kind}")
    stem = PurePath(base).stem or "out"
    commands = _rar_commands(base, f"{stem}_展開") if kind == "rar" else _image_commands(base, f"{stem}.png", kind == "avif")
    d = {"format": FORMAT[kind], "why": WHY[kind], "next_step": {"commands": commands, "after": AFTER, "writes": WRITES, "hint": HINT}}
    if kind == "rar":
        d["message"] = RAR_MESSAGE
    return d


def message(kind: str, name: str = "") -> str:
    """エラーの文に使う、1 つの文字列（形式の名前・理由・手順）。"""
    d = needs_tool(kind, name)
    cmds = "／".join(f"{c['os']}：{c['command']}" for c in d["next_step"]["commands"][:4])
    head = f"{RAR_MESSAGE}。" if kind == "rar" else f"この画像（{d['format']}）は別の道具で PNG に変換してください。"
    return f"{head}{d['why']}。例：{cmds}。{AFTER}。{WRITES}"


def build(blob: bytes, name: str) -> dict:
    kind = kind_of_blob(blob)
    if kind is None:
        raise ValueError("RAR・HEIC・AVIF のどれでもありません")
    d = needs_tool(kind, name)
    lines = [f"- {c['os']}：`{c['command']}`" + (f"（{c['note']}）" if c["note"] else "") for c in d["next_step"]["commands"]]
    secs = [("なぜ開かないか", [d["message"], d["why"]] if kind == "rar" else [d["why"]]),
            ("次の一手（入っている道具で）", [HINT, *lines]),
            ("できたファイルで", [AFTER, WRITES])]
    return rd.make(kind, KIND_JA[kind], name, blob, secs, "中身は読んでいません。形式を見分け、別の道具で変換・展開する手順だけを示します。", needs_tool=d, format=FORMAT[kind])


def grade(blob: bytes, name: str):
    kind = kind_of_blob(blob)
    win = build(blob, name)
    return rd.grade(win, "E", "別の道具で変換してから読む", message(kind, name), f"{FORMAT[kind]} には書かない", WRITE)


def audit(blob: bytes, name: str):
    from ..xlsx.audit import Audit, Check

    kind = kind_of_blob(blob)
    au = Audit(name)
    au.checks.append(Check("形式の見分け", True, f"{FORMAT[kind]} と見分けました。中身は読んでいません（{WHY[kind]}）"))
    au.checks.append(Check("別の道具で変換", True, message(kind, name), agent=True))
    return au
