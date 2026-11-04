"""PDF：印刷された字（選択肢の「有」「2.更新」「昭和」など）を、○で囲む。

指定（fill の値）：{"@c1": {"circle": 囲む字, "near": 近くの見出しの字, "page": ページ, "occurrence": 何番目か,
"pad": 字と○のすき間（pt）, "box": [左, 上, 右, 下]（mm。○の外枠をエージェントが決めるとき）, "style": "ellipse"（既定）か "border"}}。
窓は印刷された選択肢を欄として並べない（字の意味からは決めない）ので、エージェントが窓の字を読んで、囲む字を指定する。
- 囲む字は、窓の字のとおりに（字の間の空白は、あってもなくてもよい）。同じ字がいくつもあるときは、near（いちばん近い所）か
  occurrence（読む順）で1つに決める。near があれば、occurrence は near といちばん近い関係の候補の中の何番目か（Word・Excel と
  同じ）。near がなければ、ページ（page がなければ文書）のすべての候補の中の何番目か。決まらなければ、候補を添えて断る
  （一覧の番号は、そのときの指定で occurrence に渡す番号）。occurrence で指した字は、字の並びの中の字でも囲む（結果に添える）。
  near の字が何か所にもあり、それぞれに近い候補が違えば、occurrence のないときは断る
- ○の形：囲む字の箱（字の送りの幅 × 字の高さ）を合わせた四角を、すき間（pad、既定は字の大きさの 0.12 倍）だけ広げた四角の
  4 つの角を通る楕円のうち、ほかの字の面に触れないもので、上下と左右のはみ出しがいちばん近いもの。触れるときは、縦横の比を
  細かく変え、すき間を 0（字の箱の角を通る）まで小さくしてためす。ほかの字は、字の箱の縁から字の大きさの 1/12 だけ内側（字の面。
  和文の字は字の箱いっぱいには描かれない。全角の約物は JIS X 4051 の字の描かれる半分）を、触れてはいけない所とする
  （ほかの字の箱の余白には、かかってよい）
- style "border"：字の箱に沿った四角の囲み線（すき間は同じ決め方）。楕円がどれもほかの字に触れる、詰まった所のため
- どの形もほかの字に触れる・囲む字がページの外にかかるときは、書かずに理由を返す（box で外枠を決めるか、本人に伝える）
- 読み戻しは、書くときと同じ決め方（encloses・hits）：書いた後のファイルの○が、言った字の箱をすっかり囲み、ほかの字の面に触れないか

書き方：ページの描画（/Contents）の前に「q」だけのストリームを、後ろに○を描くストリームを足す（増分の保存）。後ろの
ストリームは、元の描画の q の残り（閉じていない q）を Q で戻してから、ページの初めの座標で描く。元のストリームは1バイトも
変えない。○は印（/Artifact、TdkCircle の名前つき）の中に描くので、証明（proof.py）と読み戻しで、どれを足したかが分かる。
四角の囲み線も曲線（まっすぐな 3 次ベジエ曲線）で描くので、窓は罫線（マス）と読まない。
"""

from __future__ import annotations

import math
import bisect
import re
import unicodedata

from .anchor import AnchorError
from .content import page_content, tokens
from .layout import PageLayout, View

K = 0.5522847498  # 4 つの 3 次ベジエ曲線で楕円を描くときの、制御点の比
FACE = 1 / 12  # 字の箱の縁から、字の面までの内側の幅（字の大きさに対する比）
# 字の墨（はっきり字を切る所）：字の箱の縁から、字の大きさの 0.2 倍だけ内側。字の面（FACE）には触れても、ここに触れなければ、
# ○の線は字の余白の縁をかすめるだけ（断らずに、位置を添えて知らせる。いまのエージェントは書いた後のページを見て確かめる）
INK = 0.2
TAG = "TdkCircle"
PAD = 0.12  # 字と○のすき間の既定（字の大きさに対する比）。詰まった所では 0 まで小さくする
TOL = 0.05  # 囲んでいるかを比べるときの余裕（pt）：描いた座標の丸め（小数 3 桁）と、ベジエ曲線で楕円を描く近似の分


def _lines(lay: PageLayout):
    for c in lay.cells:
        for ln in c.lines:
            yield c, ln
    for ln in lay.free_lines:
        yield None, ln


def lines_between(lay: PageLayout, top: float, bottom: float) -> list:
    """(マス, 行) のうち、行の上下が top〜bottom の中にあるもの（並びは _lines のとおり）。欄ごとに、ページの全部の行を見ない。"""
    idx = lay.__dict__.get("_lines_by_top")
    if idx is None:
        rows = sorted(((ln.y0, i, c, ln) for i, (c, ln) in enumerate(_lines(lay))), key=lambda e: (e[0], e[1]))
        idx = lay.__dict__["_lines_by_top"] = ([e[0] for e in rows], rows)
    tops, rows = idx
    a, b = bisect.bisect_left(tops, top), bisect.bisect_right(tops, bottom)
    return [(e[2], e[3]) for e in sorted((e for e in rows[a:b] if e[3].y1 <= bottom), key=lambda e: e[1])]


def _pattern(text: str) -> re.Pattern:
    """字の間の空白（全角・半角）は、あってもなくてもよい。"""
    chars = [re.escape(ch) for ch in text if not ch.isspace()]
    return re.compile(r"[ 　]*".join(chars))


def _gap(r, s) -> float:
    dx = max(0.0, max(r[0], s[0]) - min(r[2], s[2]))
    dy = max(0.0, max(r[1], s[1]) - min(r[3], s[3]))
    return math.hypot(dx, dy)


def _rank(r, s) -> tuple[int, float]:
    """目印 s から字 r への近さ：同じ行の高さ（上下が重なる）にあるものを先に、次に同じ列（左右が重なる）、ほかは距離で。"""
    if min(r[3], s[3]) - max(r[1], s[1]) > 0:
        return 0, max(0.0, max(r[0], s[0]) - min(r[2], s[2]))
    if min(r[2], s[2]) - max(r[0], s[0]) > 0:
        return 1, max(0.0, max(r[1], s[1]) - min(r[3], s[3]))
    return 2, _gap(r, s)


def _hits(lay: PageLayout, text: str):
    pat = _pattern(text)
    out = []
    for cell, ln in _lines(lay):
        for m in pat.finditer(ln.text):
            if m.end() == m.start():
                continue
            out.append((cell, ln, m.start(), m.end()))
    if not out:
        # 字のとおりに見つからなければ、正規化（NFC・NFKC）した形どうしで（分解された「e + ̂」と「ê」など）
        from .anchor import find_spans

        for cell, ln in _lines(lay):
            for a, b in find_spans(ln.text, "".join(ch for ch in text if not ch.isspace())):
                out.append((cell, ln, a, b))
    return out


def _sep(ch: str) -> bool:
    """字の区切り：字（文字・数字・結合の字。Unicode の一般カテゴリ L*・N*・M*）でないもの。空白（Z*）・区切りの記号（P*：
    「・」「、」「（」「）」「／」）・記号（S*：選択肢の前の「□」「○」）など。字の並びの中に埋まった所は、前か後ろが字。"""
    return ch.isspace() or unicodedata.category(ch)[0] not in ("L", "N", "M")


def bounded(text: str, a: int, b: int) -> bool:
    """text[a:b] が、前後を区切り（行の端・空白・区切りの記号）ではさまれているか。「有・無」の「無」ははさまれ、
    見出し「有無」の「無」（字の並びの中に埋まった字）ははさまれない。"""
    return (a == 0 or _sep(text[a - 1])) and (b >= len(text) or _sep(text[b]))


def _alone(h) -> bool:
    """候補 h（ページ, マス, 行, 始め, 終わり）が、行にその字だけがある所か（前後に、空白のほかの字がない）。"""
    t = h[2].text
    return not t[:h[3]].strip() and not t[h[4]:].strip()


def around(text: str, a: int, b: int, n: int = 10) -> str:
    """囲む字の前後の字（窓の字のまま。囲む字は【】でくくる）：どの字を囲んだかを、本人とエージェントに伝えるため。"""
    t = re.sub(r"[ 　]+", " ", text[max(0, a - n):a]).lstrip() + "【" + text[a:b] + "】" + \
        re.sub(r"[ 　]+", " ", text[b:b + n]).rstrip()
    return t


def _listing(order, show, all_hits=None, where=None) -> str:
    """候補の一覧。番号は、いまの指定（near のあり・なし）で occurrence に渡す番号（order の中の何番目か）。order にない候補
    （near からは選べない所）は、near を外したときの occurrence の番号（all_hits の中の何番目か）を添える。前後の字と、
    字の並びの中か、位置（where があれば：ページの左上からの mm）を添える。"""
    items = []
    for h in show:
        k = next((i for i, o in enumerate(order, 1) if o is h), None)
        if k is not None:
            head = f"{k}番目"
        else:
            g = next((i for i, o in enumerate(all_hits or [], 1) if o is h), None)
            head = f"（near を外して occurrence {g}）" if g else "（この指定では選べない所）"
        pos = f"・{where(h)}" if where is not None else ""
        items.append(f"{head}：{h[0]}ページ{pos}「{around(h[2].text, h[3], h[4])}」"
                     + ("" if bounded(h[2].text, h[3], h[4]) else "（字の並びの中）"))
    return "／".join(items[:10])


def _on_near_line(h, an) -> bool:
    """候補 h が、near の字の行の高さにあるか（上下が重なる。near の字のマスではなく、near の字そのものの行で比べる）。"""
    if an[0] != h[0]:
        return False
    if h[2] is an[3]:
        return True
    r, s_ = h[2].span_rect(h[3], h[4]), an[3].span_rect(an[4], an[5])
    return min(r[3], s_[3]) - max(r[1], s_[1]) > 0


def find(doc, spec: dict, layouts: dict, info: dict | None = None) -> tuple[int, object, object, int, int]:
    """囲む字の所：(ページ, マス, 行, 始め, 終わり)。info：near の字の行にない候補を選んだとき、知らせること（off_line）。"""
    target = spec.get("circle")
    if not isinstance(target, str) or not target.strip():
        raise AnchorError('○で囲むときは {"circle": 囲む字, "near": 近くの見出し（省ける）, "page": ページ（省ける）} の形で指定してください')
    pages = doc.pages()
    page_no = spec.get("page")
    if page_no is not None and not (isinstance(page_no, int) and not isinstance(page_no, bool) and 1 <= page_no <= len(pages)):
        raise AnchorError(f"page は 1〜{len(pages)} で指定してください")
    hits = []
    for n in ([page_no] if page_no else range(1, len(pages) + 1)):
        lay = layout(doc, n, layouts)
        for cell, ln, a, b in _hits(lay, target):
            hits.append((n, cell, ln, a, b))
    hits.sort(key=lambda h: (h[0], round(h[2].y0, 1), round(h[2].x0, 1), h[3]))
    within = spec.get("within")
    if within is not None:
        # 窓の選択肢の所（見た目の座標の矩形）：その中に字の中心があるもの
        if not (isinstance(within, list) and len(within) == 4 and page_no):
            raise AnchorError("within は page と [左, 上, 右, 下]（pt）で指定してください")
        def _in(h):
            r = h[2].span_rect(h[3], h[4])
            cx, cy = (r[0] + r[2]) / 2, (r[1] + r[3]) / 2
            return within[0] - 0.5 <= cx <= within[2] + 0.5 and within[1] - 0.5 <= cy <= within[3] + 0.5
        hits = [h for h in hits if _in(h)]
        if len(hits) != 1:
            raise AnchorError(f"選択肢「{target}」の字の所が、ページの字から1つに決まりません")
        return hits[0]
    where = f" {page_no} ページに" if page_no else ""
    if not hits:
        raise AnchorError(f"「{target}」が{where}見つかりません（窓の字のとおりに指定してください）")
    occ = spec.get("occurrence")
    near = spec.get("near")
    has_near = isinstance(near, str) and bool(near.strip())
    if occ is not None and (not isinstance(occ, int) or isinstance(occ, bool) or occ < 1):
        raise AnchorError("occurrence は 1 以上の整数で指定してください")

    def where_mm(h) -> str:
        lay_ = layout(doc, h[0], layouts)
        r = h[2].span_rect(h[3], h[4])
        return f"左 {r[0] / lay_.view.mm:.0f}・上 {r[1] / lay_.view.mm:.0f} mm"

    # 選ぶ候補は、字の形だけで決める（字の意味は見ない）：
    # 1) near の字そのものの中にある所（見出し「…の有無」の「無」）は、けっして選ばない
    # 2) occurrence があれば、その番号の候補（字の並びの中の字でもよい：エージェントが番号で指したもの。結果にそう添える）。
    #    番号は、near がなければページ（page がなければ文書）のすべての候補の読む順、near があれば near といちばん近い関係
    #    （同じマス・同じ行…：_rank の初めの数）の候補の読む順。断るときの一覧も、同じ番号で示す
    # 3) occurrence がなければ、前後が区切り（空白・区切りの記号 Z*・P*・行の端）ではさまれた所（「有・無」の「無」）を、字の
    #    並びの中に埋まった所（見出し「有無」の「無」）より先にする。埋まった所しかなければ、選ばずに断る
    anchors = []
    if has_near:
        for n in sorted({h[0] for h in hits}):
            lay = layout(doc, n, layouts)
            for cell, ln, a, b in _hits(lay, near):
                # 見出しが表のマスの中なら、近さはそのマスの外枠で測る（縦書きの「変更内容」のマスの行の高さにある選択肢を、
                # マスの外の下の行の見出しより先にする）。マスの外の字は、その字の所で測る
                r = (cell.x0, cell.y0, cell.x1, cell.y1) if cell is not None else ln.span_rect(a, b)
                anchors.append((n, r, cell, ln, a, b))
        if not anchors:
            raise AnchorError(f"near の「{near}」が、「{target}」のあるページに見つかりません")
    all_hits = hits  # near のないときの occurrence の番号は、この読む順
    order = hits

    def rank(h, ans):
        return min((_rank(h[2].span_rect(h[3], h[4]), an[1]) for an in ans if an[0] == h[0]), default=(9, float("inf")))

    inside: list = []  # near の字の中にある候補（選択肢の項目でないもの）と、その near の所
    if anchors:
        from ..circlepick import inside_message, items_message, option_item

        def near_of(h):
            return next((an for an in anchors if an[0] == h[0] and an[3] is h[2] and an[4] <= h[3] and h[4] <= an[5]), None)

        selfs = [(h, near_of(h)) for h in hits]
        items = [h for h, an in selfs if an is not None and option_item(h[2].text, h[3], h[4], an[4], an[5])]
        if items:
            # near が選択肢の並び（「大・昭・平・令」）で、囲む字がその項目：その所を囲む
            if occ is not None:
                if occ > len(items):
                    raise AnchorError(f"near の「{near}」の中の「{target}」は {len(items)} か所です"
                                      f"（{_listing(items, items, all_hits, where_mm)}）。occurrence は 1〜{len(items)} で指定してください")
                return items[occ - 1]
            if len(items) == 1:
                return items[0]
            raise AnchorError(items_message(target, near, len(items), _listing(items, items, all_hits, where_mm)))
        inside = [(h, an) for h, an in selfs if an is not None]
        hits = [h for h, an in selfs if an is None]
        if not hits:
            if inside and occ is not None and occ <= len(inside):
                return inside[occ - 1][0]
            raise AnchorError(f"「{target}」は、near の「{near}」の字の中にしか見つかりません（区切られた選択肢の項目では"
                              "ありません）。その字を囲むなら occurrence で指定してください："
                              f"{_listing([h for h, _a in inside], [h for h, _a in inside], all_hits, where_mm)}")
        ranks = [rank(h, anchors) for h in hits]
        best_kind = min(r[0] for r in ranks)
        order = [h for h, r in zip(hits, ranks) if r[0] == best_kind] + [h for h, _a in inside]
    if occ is not None:
        if occ > len(order):
            if anchors:
                raise AnchorError(f"near の「{near}」に近い「{target}」は {len(order)} か所です（{_listing(order, order, all_hits, where_mm)}）。"
                                  f"occurrence は、その中の何番目か（1〜{len(order)}）で指定してください")
            raise AnchorError(f"occurrence は 1〜{len(order)} で指定してください（「{target}」は{where} {len(order)} か所）")
        return order[occ - 1]
    # near の字の行にある候補があれば、その中から選ぶ（マスの外枠で近くても、別の行の字に移らない。同じくらい近ければ断る）
    online = [h for h in hits if any(_on_near_line(h, an) for an in anchors)]
    got = _pick_free(online or hits, anchors, order, all_hits, where_mm, rank, near, target)
    if inside and not any(_on_near_line(got, an) for _h, an in inside):
        # 囲む字は near の字の中にもある（「都県市」の「都」）。いちばん近いほかの候補は別の行：黙って移らずに断る
        raise AnchorError(inside_message(target, near, _listing(order, order, all_hits, where_mm)))
    if anchors and info is not None and not any(_on_near_line(got, an) for an in anchors):
        an = min((x for x in anchors if x[0] == got[0]), key=lambda x: _gap(got[2].span_rect(got[3], got[4]), x[3].span_rect(x[4], x[5])))
        info["off_line"] = (f"囲んだ「{target}」は、near の「{near}」の行にありません。囲んだ所：{where_mm(got)} の行"
                            f"「{around(got[2].text, got[3], got[4])}」。near の行：{where_mm((an[0], None, an[3], an[4], an[5]))} の"
                            f"「{around(an[3].text, an[4], an[5])}」。別の字のことがあるので、窓と view_page で確かめてください"
                            "（ほかの候補は occurrence で選べます）")
    return got


def _pick_free(hits, anchors, order, all_hits, where_mm, rank, near, target):
    """occurrence のないときの選び方（near の字の中の候補は、hits から除いてある）。"""
    if len(anchors) > 1:
        # near の字が何か所にもある（「浴」が「入浴」にも「浴槽」にもある）：それぞれのいちばん近い候補が違えば、どの near か
        # 決まらないので断る（黙って片方を選ばない）
        picks: dict = {}
        for an in anchors:
            cand = [h for h in hits if h[0] == an[0] and bounded(h[2].text, h[3], h[4])] or [h for h in hits if h[0] == an[0]]
            if cand:
                best_h = min(cand, key=lambda h: rank(h, [an]))
                picks.setdefault(id(best_h), (an, best_h))
        if len(picks) > 1:
            places = "／".join(f"{an[0]}ページ「{around(an[3].text, an[4], an[5])}」" for an, _h in list(picks.values())[:6])
            raise AnchorError(f"near の「{near}」が {len(anchors)} か所（{places}）にあり、それぞれに近い「{target}」が違うので、"
                              "どれを囲むか決まりません。near を、ほかの所にない字（見出しの全体など）にするか、occurrence で"
                              f"指定してください：{_listing(order, [h for _a, h in picks.values()], all_hits, where_mm)}")
    free = [h for h in hits if bounded(h[2].text, h[3], h[4])]
    if not free:
        raise AnchorError(f"「{target}」は、どれもほかの字の並びの中にあります（区切りの字ではさまれた所がありません）。"
                          f"選択肢の字か確かめて、囲むなら occurrence（何番目か）で指定してください："
                          f"{_listing(order, hits + [h for h in order if not any(h is x for x in hits)], all_hits, where_mm)}")
    if anchors and len(free) > 1:
        scored = sorted(((rank(h, anchors), h) for h in free), key=lambda x: x[0])
        best = scored[0]
        tol = 0.25 * best[1][2].size
        tied = [h for d, h in scored if d[0] == best[0][0] and d[1] - best[0][1] <= tol]
        if len(tied) > 1:
            raise AnchorError(f"「{near}」に同じくらい近い「{target}」が、2 か所以上あります。occurrence で1つに決めてください："
                              f"{_listing(order, tied, all_hits, where_mm)}")
        listed = [h for h in free if not _alone(h)]
        if _alone(best[1]) and listed:
            # いちばん近いのが、行にその字しかない所（見出し「変更後 振込口座」のことがある）で、ほかに、行の中に区切って
            # 並んだ所（選択肢の並び「３ 送付先　４ 振込口座」）がある：どちらを囲むかは形からは決めず、断る
            raise AnchorError(f"いちばん近い「{target}」は、行にその字だけがある所（見出しのことがある）で、ほかに、行の中に"
                              f"区切って並んだ「{target}」もあります。どれを囲むか、occurrence で1つに決めてください："
                              f"{_listing(order, free, all_hits, where_mm)}")
        return best[1]
    if len(free) > 1:
        raise AnchorError(f"「{target}」が {len(free)} か所にあります。near（近くの見出しの字）か occurrence（何番目か）で"
                          f"1つに決めてください：{_listing(order, free, all_hits, where_mm)}")
    return free[0]


def layout(doc, n: int, layouts: dict) -> PageLayout:
    if n not in layouts:
        pg = doc.pages()[n - 1]
        layouts[n] = PageLayout(doc, pg, page_content(doc, pg), n)
    return layouts[n]


def target_chars(lay: PageLayout, ln, a: int, b: int) -> list:
    """行の text[a:b] に当たる、ページの字（空白でないもの）。"""
    r = ln.span_rect(a, b)
    out = []
    for c in ln.chars:
        if c.text.isspace():
            continue
        if ln.vertical:
            if r[1] - 0.01 <= c.cy <= r[3] + 0.01:
                out.append(c)
        elif r[0] - 0.01 <= c.cx <= r[2] + 0.01:
            out.append(c)
    return out


# JIS X 4051（日本語文書の組版方法）の約物の字の面：全角の箱のうち、字が描かれるのは半分（二分）。横書きで、
# 始め括弧類は右半分、終わり括弧類・句点類・読点類は左半分、中点類は中央の半分
OPEN_BRACKETS = set("‘“（〔［｛〈《「『【⦅〘〖«〝")
CLOSE_BRACKETS = set("’”）〕］｝〉》」』】⦆〙〗»〟")
FULL_STOPS = set("。．、，")
MIDDLE_DOTS = set("・：；")


def face(c, inset: float = FACE, strict: bool = False) -> tuple[float, float, float, float]:
    """字の面：字の箱の縁から、字の大きさの 1/12（inset）だけ内側。全角の約物（JIS X 4051 の括弧類・句読点）は、字の描かれる
    半分の箱で、中点類は箱の中央の半分で（横書きの字のみ。半角の箱の約物は、箱がもう字の描かれる半分）。
    strict：約物（Unicode の一般カテゴリ P*）は、その字の描かれる箱の左右の端まで字の墨があるものとして、内側に入れない
    （○の線が「・」の点や「（」のふくらみを通るのを、知らせるため）。"""
    m = inset * c.size
    mx = 0.0 if strict and any(unicodedata.category(ch)[0] == "P" for ch in c.text) else m
    bx0, bx1 = c.x0, c.x1
    if len(c.text) == 1 and not c.vertical:
        half = (bx1 - bx0) / 2
        if c.text in MIDDLE_DOTS:
            bx0, bx1 = bx0 + half / 2, bx1 - half / 2
        elif bx1 - bx0 > 0.75 * c.size:
            if c.text in OPEN_BRACKETS:
                bx0 = bx0 + half
            elif c.text in CLOSE_BRACKETS or c.text in FULL_STOPS:
                bx1 = bx1 - half
    x0, y0, x1, y1 = bx0 + mx, c.y0 + m, bx1 - mx, c.y1 - m
    if x1 <= x0:
        x0 = x1 = (bx0 + bx1) / 2
    if y1 <= y0:
        y0 = y1 = (c.y0 + c.y1) / 2
    return x0, y0, x1, y1


def touches(e: tuple[float, float, float, float], box) -> bool:
    """楕円（中心 cx, cy・半径 a, b）の内側（縁を含む）と、四角が重なるか。"""
    cx, cy, a, b = e
    # 楕円を単位円にする縮め方で、四角も縮める（軸にそろった四角のまま）
    x0, y0, x1, y1 = (box[0] - cx) / a, (box[1] - cy) / b, (box[2] - cx) / a, (box[3] - cy) / b
    dx = 0.0 if x0 <= 0 <= x1 else min(abs(x0), abs(x1))
    dy = 0.0 if y0 <= 0 <= y1 else min(abs(y0), abs(y1))
    return dx * dx + dy * dy <= 1.0


def inside(e, box) -> bool:
    """四角が、楕円の内側にすっかり入るか（4 つの角が内側）。"""
    cx, cy, a, b = e
    return all(((x - cx) / a) ** 2 + ((y - cy) / b) ** 2 <= 1.0 + 1e-9 for x in (box[0], box[2]) for y in (box[1], box[3]))


def _rect_meets(r, box) -> bool:
    """四角どうしが重なるか（縁を含む）。"""
    return r[0] <= box[2] and box[0] <= r[2] and r[1] <= box[3] and box[1] <= r[3]


def encloses(kind: str, g, c, lw: float = 0.0) -> bool:
    """囲みの形（kind "ellipse" は楕円 (cx, cy, a, b)、"border" は四角 (x0, y0, x1, y1)。見た目の座標）が、字 c の箱をすっかり
    囲むか。書くとき（resolve）と読み戻し（read_back）・窓の「囲まれた字」で同じ決め方：描いた座標の丸め（小数 3 桁）とベジエ曲線の
    近似の分（TOL）だけ広げて比べる。lw（線の太さ）があれば、線の内側の縁（太さの半分だけ内側）が字の箱をすっかり囲むか
    （線が字の箱にかからない。字の箱の角まで字が描かれていることがあるので、箱の中心でなく箱そのもので比べる）。"""
    box = (c.x0, c.y0, c.x1, c.y1)
    h = lw / 2
    if kind == "border":
        return g[0] + h - TOL <= box[0] and g[1] + h - TOL <= box[1] and box[2] <= g[2] - h + TOL and box[3] <= g[3] - h + TOL
    if g[2] - h <= 0 or g[3] - h <= 0:
        return False
    # 楕円の線の内側の縁は、軸を太さの半分だけ縮めた楕円を含む（縮めた楕円で比べれば、線は字の箱にかからない）
    return inside((g[0], g[1], g[2] - h + TOL, g[3] - h + TOL), box)


def _outer(kind: str, g, lw: float):
    """囲みの線の外側の縁（線の太さの半分だけ外へ）。"""
    h = lw / 2
    if kind == "border":
        return (g[0] - h, g[1] - h, g[2] + h, g[3] + h)
    return (g[0], g[1], g[2] + h, g[3] + h)


def hits(kind: str, g, c, lw: float = 0.0, loose: bool = False, strict: bool = False) -> bool:
    """囲みの線（太さ lw）が、字 c の面（face）に触れるか（ほかの字に触れてはいけない所）。書くときと読み戻しで同じ決め方。
    loose：字の墨（INK）に触れるか（字の面をかすめるだけの○を、知らせて描いたもの）。strict：face の strict（知らせるとき）。"""
    f = face(c, INK if loose else FACE, strict)
    o = _outer(kind, g, lw)
    if kind == "border":
        return _rect_meets(o, f)
    return touches(o, f)


def ellipses_meet(e1, e2) -> bool:
    """2 つの楕円が重なるか（片方の縁の点が、もう片方の内側にある）。"""
    for e, f in ((e1, e2), (e2, e1)):
        cx, cy, a, b = e
        for k in range(72):
            t = 2 * math.pi * k / 72
            x, y = cx + a * math.cos(t), cy + b * math.sin(t)
            if ((x - f[0]) / f[2]) ** 2 + ((y - f[1]) / f[3]) ** 2 <= 1.0:
                return True
    return False


def shapes_meet(k1: str, g1, k2: str, g2, lw: float = 0.0) -> bool:
    """2 つの囲み（楕円・四角）が重なるか（線の太さ lw の半分ずつ外へ広げて）。四角どうしは四角のまま比べる（外接する楕円で
    比べると、6 mm 離れた隣の行の四角どうしも重なって見え、断ってしまう）。四角と楕円は、楕円の内側と四角が重なるか。"""
    h = lw / 2
    if k1 == "border" and k2 == "border":
        return _rect_meets((g1[0] - h, g1[1] - h, g1[2] + h, g1[3] + h), (g2[0] - h, g2[1] - h, g2[2] + h, g2[3] + h))
    if k1 == "border" or k2 == "border":
        r, e = (g1, g2) if k1 == "border" else (g2, g1)
        return touches((e[0], e[1], e[2] + h, e[3] + h), (r[0] - h, r[1] - h, r[2] + h, r[3] + h))
    return ellipses_meet((g1[0], g1[1], g1[2] + h, g1[3] + h), (g2[0], g2[1], g2[2] + h, g2[3] + h))


def _pads(size: float, pad: float | None) -> list[float]:
    """ためすすき間（pt）：指定があればそれだけ、なければ字の大きさの 0.12 倍から、詰まった所では 0 まで小さくする。"""
    if pad is not None:
        return [pad]
    return [PAD * size * k for k in (1.0, 2 / 3, 1 / 3)] + [0.0]


def shape(rect, size: float, others: list, page: tuple[float, float] | None, pad: float | None = None,
          targets: list | None = None, loose: bool = False, avoid=None, strict: bool = False):
    """囲む字の四角 rect（字の箱を合わせたもの）を囲む楕円 (cx, cy, a, b)：rect をすき間だけ広げた四角の角を通る楕円
    （縦横の比を細かく変えて）のうち、ほかの字の面に触れないもの。すき間は大きいほうから、0（字の箱の角を通る）まで小さくし、
    同じすき間では、左右と上下のはみ出しがいちばん近いものを選ぶ。見つからなければ (None, いちばん多く触れた字)。"""
    cx, cy = (rect[0] + rect[2]) / 2, (rect[1] + rect[3]) / 2
    blockers: dict[int, list] = {}
    near = [(o, face(o, INK if loose else FACE, strict)) for o in others]
    h_lw = line_width(size) / 2  # 線の太さの半分：線の外側の縁が、ほかの字の面に触れない
    # 縦横の比：角を通る楕円 (w/2a)^2 = t、(h/2b)^2 = 1 - t。t が 0・1 に近い（細長い）所は細かく
    ts = sorted({k / 200 for k in range(1, 200)} | {1 - 0.5 ** k for k in range(8, 15)} | {0.5 ** k for k in range(8, 15)})
    for m in _pads(size, pad):
        # 線の内側の縁が、囲む字の箱にかからないよう、すき間に線の太さの半分を足す（すき間 0 でも、線は字の箱の外）
        w, h = rect[2] - rect[0] + 2 * (m + h_lw), rect[3] - rect[1] + 2 * (m + h_lw)
        best = None
        for t in ts:
            a, b = (w / 2) / math.sqrt(t), (h / 2) / math.sqrt(1 - t)
            if a - w / 2 > 1.5 * size or b - h / 2 > 1.5 * size:
                continue  # はみ出しが字の大きさの 1.5 倍を超える楕円は、どこを囲んだか分かりにくい
            e = (cx, cy, a, b)
            if page is not None and (cx - a < 0 or cy - b < 0 or cx + a > page[0] or cy + b > page[1]):
                continue
            if targets and not all(encloses("ellipse", e, c, 2 * h_lw) for c in targets):
                continue  # 線が、囲む字の箱にかかる（字を切って見える）
            if avoid is not None and avoid("ellipse", e):
                continue  # ページの罫線を横切る（線のあいだに収まる、平たい・小さい楕円を先にためす）
            hit = next((o for o, f in near if touches((cx, cy, a + h_lw, b + h_lw), f)), None)
            if hit is not None:
                blockers.setdefault(id(hit), [hit, 0])[1] += 1
                continue
            score = abs((a - w / 2) - (b - h / 2))
            if best is None or score < best[0]:
                best = (score, e)
        if best:
            return best[1], None
    return None, max(blockers.values(), key=lambda x: x[1])[0] if blockers else None


def border(rect, size: float, others: list, page: tuple[float, float] | None, pad: float | None = None,
           loose: bool = False, avoid=None, strict: bool = False):
    """囲む字の四角 rect を、すき間だけ広げた四角の囲み線（style "border"）：ほかの字の面に触れない、いちばん大きなすき間のもの。"""
    blocker = None
    h_lw = line_width(size) / 2  # 線の内側の縁が、囲む字の箱にかからないよう、すき間に線の太さの半分を足す
    for m in _pads(size, pad):
        r = (rect[0] - m - h_lw, rect[1] - m - h_lw, rect[2] + m + h_lw, rect[3] + m + h_lw)
        if page is not None and (r[0] < 0 or r[1] < 0 or r[2] > page[0] or r[3] > page[1]):
            continue
        if avoid is not None and avoid("border", r):
            continue
        hit = next((o for o in others if hits("border", r, o, line_width(size), loose)), None)
        if hit is None:
            return r, None
        blocker = blocker or hit
    return None, blocker


def resolve(doc, spec: dict, layouts: dict, used: dict, taken: set | None = None) -> dict:
    """○で囲む所を決める（書き込みと証明で同じ決め方）。返す：{id, page, text, style, ellipse（楕円のとき。見た目の座標）,
    border（四角の囲み線のとき）, label, context, size}。"""
    style = spec.get("style", "ellipse")
    if style not in ("ellipse", "border"):
        raise AnchorError('style は "ellipse"（楕円の○。既定）か "border"（字を囲む四角の囲み線）です')
    info: dict = {}
    try:
        n, cell, ln, a, b = find(doc, spec, layouts, info)
    except AnchorError:
        if spec.get("box") is not None and spec.get("page"):
            # 枠（box）を決めた○は、ページの字が読めなくても描ける（画像だけのページの選択肢など）。囲んだのが言った字かは
            # 確かめられないので、そう伝える
            return _resolve_unread_box(doc, spec, layouts, used, taken, style)
        raise
    lay = layout(doc, n, layouts)
    mine = target_chars(lay, ln, a, b)
    if not mine:
        raise AnchorError(f"「{spec['circle']}」の字の位置が分かりません")
    ids = {id(c) for c in mine}
    rect = (min(c.x0 for c in mine), min(c.y0 for c in mine), max(c.x1 for c in mine), max(c.y1 for c in mine))
    size = max(c.size for c in mine)
    others = [c for c in lay.chars if id(c) not in ids and not c.text.isspace()
              and c.x1 > rect[0] - 4 * size and c.x0 < rect[2] + 4 * size and c.y1 > rect[1] - 4 * size and c.y0 < rect[3] + 4 * size]
    pad = spec.get("pad")
    if pad is not None and not (isinstance(pad, (int, float)) and not isinstance(pad, bool) and 0 <= pad <= 3 * size):
        raise AnchorError(f"pad は、字と○のすき間（pt、0〜{3 * size:.0f}）で指定してください")
    page = (lay.view.width, lay.view.height)
    box = spec.get("box")
    if box is not None:
        if not (isinstance(box, list) and len(box) == 4 and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in box)):
            raise AnchorError("box は [左, 上, 右, 下]（ページの左上からの mm。○の外枠）で指定してください")
        x0, y0, x1, y1 = (v * lay.view.mm for v in box)  # mm → ページの座標（UserUnit も反映）
        if not (0 <= x0 < x1 <= page[0] + 0.01 and 0 <= y0 < y1 <= page[1] + 0.01):
            raise AnchorError("box がページの外か、左＜右・上＜下になっていません")
        g = (x0, y0, x1, y1) if style == "border" else ((x0 + x1) / 2, (y0 + y1) / 2, (x1 - x0) / 2, (y1 - y0) / 2)
        if not all(encloses(style, g, c, line_width(size)) for c in mine):
            raise AnchorError(f"box の○（線の内側の縁）が、「{spec['circle']}」の字を囲みきれません（字の箱は 左 {rect[0] / lay.view.mm:.1f}・上 "
                              f"{rect[1] / lay.view.mm:.1f}・右 {rect[2] / lay.view.mm:.1f}・下 {rect[3] / lay.view.mm:.1f} mm）")
        loose = None
        if any(hits(style, g, o, line_width(size)) for o in others):
            hit = next((o for o in others if hits(style, g, o, line_width(size), True)), None)
            if hit is not None:
                mm = lay.view.mm
                raise AnchorError(f"box の○が、ほかの字「{hit.text}」（左 {hit.x0 / mm:.1f}・上 {hit.y0 / mm:.1f}・右 {hit.x1 / mm:.1f}・"
                                  f"下 {hit.y1 / mm:.1f} mm）の字の墨を切ります")
            loose = True
    else:
        fit = border if style == "border" else (lambda *a_, **k_: shape(*a_, targets=mine, **k_))
        from .placement import _rule_crosses_shape, rule_touches_shape

        # ○の届く所（はみ出しは字の大きさの 1.5 倍まで：shape）の近くの罫線だけ比べる
        rx0, ry0, rx1, ry1 = rect[0] - 4 * size, rect[1] - 4 * size, rect[2] + 4 * size, rect[3] + 4 * size
        rules = [(r_, True) for r_ in lay.hs if ry0 <= r_.pos <= ry1 and r_.a <= rx1 and rx0 <= r_.b] + \
                [(r_, False) for r_ in lay.vs if rx0 <= r_.pos <= rx1 and r_.a <= ry1 and ry0 <= r_.b]

        def avoid(kind, g_):
            return any(_rule_crosses_shape(r_, hz, kind, g_, line_width(size)) for r_, hz in rules)

        def clear(kind, g_):
            return any(rule_touches_shape(r_, hz, kind, g_, line_width(size)) for r_, hz in rules)

        # 罫線にも約物の墨にも触れないものを先に。なければ順にゆるめる（触れるものは、描いて位置を知らせる：下と placement）
        g, hit = None, None
        for av, st in ((clear, True), (avoid, True), (avoid, False), (None, True), (None, False)):
            g, hit = fit(rect, size, others, page, pad, avoid=av, strict=st)
            if g is not None:
                break
        loose = None
        if g is None:
            # 字の面には触れる所しかない：字の墨（INK）を切らないものなら描き、触れる字と位置を知らせる
            g, hit2 = fit(rect, size, others, page, pad, loose=True)
            loose = True
            if g is None:
                other = "ellipse" if style == "border" else "border"
                ofit = border if other == "border" else (lambda *a_, **k_: shape(*a_, targets=mine, **k_))
                alt = ofit(rect, size, others, page, pad, loose=True)[0] is not None
                what = "四角で囲む" if style == "border" else "楕円で囲む"
                blk = hit2 or hit
                raise AnchorError(f"「{spec['circle']}」を{what}と、どの大きさ・縦横の比でも、ほかの字" + (f"「{blk.text}」" if blk is not None else "") +
                                  "の字の墨を切ります（字の箱どうしが重なっています）。"
                                  + (f"\"style\": \"{other}\" なら描けます。" if alt else
                                     "楕円・四角のどちらでも切るので、○は描けません。box（○の外枠、mm）で自分で決めるか、"
                                     "○の代わりに選んだ字の近くの空きに「○」の字を書く（after / box）か、本人に伝えてください"))
    for k_, kind2, g2 in used.get(n, []):
        if shapes_meet(style, g, kind2, g2, line_width(size)):
            raise AnchorError(f"「{spec['circle']}」の○が、{n} ページのほかの○と重なります")
    taken_nums = {int(m[1]) for nm in (taken or set()) for m in [re.fullmatch(rf"p{n}\.c(\d+)", nm)] if m}
    taken_nums |= {k_ for k_, _kd, _g in used.get(n, [])}
    k = next(x for x in range(1, 10 ** 6) if x not in taken_nums)
    used.setdefault(n, []).append((k, style, g))
    context = re.sub(r"[ 　]+", " ", ln.text).strip()
    label = spec.get("near") or (re.sub(r"[ 　]+", " ", cell.text()).strip()[:30] if cell is not None else context[:30])
    out = {"id": f"p{n}.c{k}", "page": n, "text": "".join(c.text for c in mine), "style": style,
           "ellipse": [round(v, 3) for v in g] if style == "ellipse" else None,
           "label": label, "context": context[:60], "size": size, "rect": [round(v, 3) for v in rect],
           "around": around(ln.text, a, b)}  # 囲んだ字の前後の字（どの「無」を囲んだかを伝える）
    ob_ = _outer(style, g, line_width(size))
    ob_ = ob_ if style == "border" else (ob_[0] - ob_[2], ob_[1] - ob_[3], ob_[0] + ob_[2], ob_[1] + ob_[3])
    out["box_mm"] = [round(v / lay.view.mm, 1) for v in ob_]  # ○の外枠（ページの左上からの mm。view_page で確かめる）
    near_c = [o for o in others if hits(style, g, o, line_width(size), strict=True)]
    if loose:
        out["loose"] = True
    notes = []
    if loose or near_c:
        mm = lay.view.mm
        ob = _outer(style, g, line_width(size))
        ob = ob if style == "border" else (ob[0] - ob[2], ob[1] - ob[3], ob[0] + ob[2], ob[1] + ob[3])
        o0 = near_c[0] if near_c else None
        at = f"。その字は 左 {o0.x0 / mm:.1f}・上 {o0.y0 / mm:.1f} mm" if o0 is not None else ""
        notes.append(f"{out['id']}：「{out['text']}」の{'四角' if style == 'border' else '○'}の線が、ほかの字"
                     f"「{''.join(o.text for o in near_c[:6])}」の縁に触れます（{n}ページ・左 {ob[0] / mm:.1f}・"
                     f"上 {ob[1] / mm:.1f}・右 {ob[2] / mm:.1f}・下 {ob[3] / mm:.1f} mm{at}）。view_page で見て確かめてください")
    if info.get("off_line"):
        notes.append(f"{out['id']}：{info['off_line']}")
    if style == "ellipse" and g[3] - (rect[3] - rect[1]) / 2 > rect[3] - rect[1]:
        # 隣の字をよけるために、字の高さより上下へ大きくはみ出した楕円：どの行を囲んだかが分かりにくい
        mm = lay.view.mm
        notes.append(f"{out['id']}：「{out['text']}」の○は、隣の字をよけて縦に長くなりました（高さ {2 * g[3] / mm:.1f} mm。字の高さは "
                     f"{(rect[3] - rect[1]) / mm:.1f} mm）。view_page で見て、上下の行にかかっていれば、\"style\": \"border\"（字に沿った四角）か box で"
                     "書き直してください")
    if notes:
        out["note"] = "\n".join(notes)
    if not bounded(ln.text, a, b):
        out["inside_word"] = True  # occurrence で指した、字の並びの中の字（「【都】道府県」）
    if style == "border":
        out["border"] = [round(v, 3) for v in g]
    return out


UNREAD_BOX = ("ページの字として読めない所（画像など）に、指定の枠（box）で描きました。囲んだのが「{t}」かは、字から確かめて"
              "いません（view_page で見て確かめてください）")


def _resolve_unread_box(doc, spec: dict, layouts: dict, used: dict, taken: set | None, style: str) -> dict:
    """囲む字がページの字として見つからないときの、枠（box・mm）で決めた○：枠がページの中で、ほかの字にも、ほかの○にも
    触れないことだけを確かめる。囲んだ字は確かめていない（unread）と返す。"""
    n = spec["page"]
    if not (isinstance(n, int) and not isinstance(n, bool) and 1 <= n <= len(doc.pages())):
        raise AnchorError(f"page は 1〜{len(doc.pages())} で指定してください")
    box = spec["box"]
    if not (isinstance(box, list) and len(box) == 4 and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in box)):
        raise AnchorError("box は [左, 上, 右, 下]（ページの左上からの mm。○の外枠）で指定してください")
    lay = layout(doc, n, layouts)
    page = (lay.view.width, lay.view.height)
    x0, y0, x1, y1 = (v * lay.view.mm for v in box)  # mm → ページの座標（UserUnit も反映）
    if not (0 <= x0 < x1 <= page[0] + 0.01 and 0 <= y0 < y1 <= page[1] + 0.01):
        raise AnchorError("box がページの外か、左＜右・上＜下になっていません")
    size = min(y1 - y0, 10.5)
    g = (x0, y0, x1, y1) if style == "border" else ((x0 + x1) / 2, (y0 + y1) / 2, (x1 - x0) / 2, (y1 - y0) / 2)
    hit = next((o for o in lay.chars if not o.text.isspace() and hits(style, g, o, line_width(size))), None)
    if hit is not None:
        raise AnchorError(f"box の○が、ページの字「{hit.text}」に触れます")
    for _k, kind2, g2 in used.get(n, []):
        if shapes_meet(style, g, kind2, g2, line_width(size)):
            raise AnchorError(f"「{spec['circle']}」の○が、{n} ページのほかの○と重なります")
    taken_nums = {int(m[1]) for nm in (taken or set()) for m in [re.fullmatch(rf"p{n}\.c(\d+)", nm)] if m}
    taken_nums |= {k_ for k_, _kd, _g in used.get(n, [])}
    k = next(x for x in range(1, 10 ** 6) if x not in taken_nums)
    used.setdefault(n, []).append((k, style, g))
    out = {"id": f"p{n}.c{k}", "page": n, "text": spec["circle"], "style": style,
           "ellipse": [round(v, 3) for v in g] if style == "ellipse" else None,
           "label": spec.get("near") or f"{n}ページの枠（左 {box[0]:g}・上 {box[1]:g} mm）", "context": "",
           "size": size, "rect": [round(v, 3) for v in (x0, y0, x1, y1)], "around": None,
           "unread": UNREAD_BOX.format(t=spec["circle"])}
    if style == "border":
        out["border"] = [round(v, 3) for v in g]
    return out


# ---------------------------------------------------------------------- 描く
def line_width(size: float) -> float:
    """○の線の太さ：字の大きさに比例（A3 の大きな字でも、字の線と同じくらいの太さに）。細すぎないよう 0.6pt 以上。"""
    return round(max(0.6, 0.07 * size), 2)


def open_depth(data: bytes) -> int:
    """描画の終わりで閉じていない q の数。"""
    d = 0
    for kind, val in tokens(data):
        if kind == "o":
            if val == "q":
                d += 1
            elif val == "Q" and d > 0:
                d -= 1
    return d


def _seg(p, q) -> str:
    """直線を 3 次ベジエ曲線で（制御点を線の 1/3・2/3 に置く）。四角の囲み線を曲線だけで描き、窓が罫線（マス）と読まないように。"""
    (x0, y0), (x1, y1) = p, q
    return (f"{x0 + (x1 - x0) / 3:.3f} {y0 + (y1 - y0) / 3:.3f} {x0 + 2 * (x1 - x0) / 3:.3f} {y0 + 2 * (y1 - y0) / 3:.3f} "
            f"{x1:.3f} {y1:.3f} c\n")


def ops(view: View, circles: list[dict], depth: int) -> bytes:
    """後ろに足すストリーム：閉じていない q を戻し（前に足した q の分も）、ページの初めの座標で楕円（四角の囲み線）を描く。"""
    out = bytearray(b"Q\n" * (depth + 1))
    for sd in circles:
        lw = line_width(sd["size"])
        if sd.get("style") == "border":
            x0, y0, x1, y1 = view.back_rect(*sd["border"])
            pts = [(x0, y0), (x1, y0), (x1, y1), (x0, y1), (x0, y0)]
            path = f"{x0:.3f} {y0:.3f} m\n" + "".join(_seg(p, q) for p, q in zip(pts, pts[1:]))
            out += (f"q\n/Artifact <</Type /Layout /{TAG} ({sd['id']}) /TdkStyle /border>> BDC\n0 G {lw} w [] 0 d 0 J 0 j\n"
                    f"{path}h S\nEMC\nQ\n").encode()
            continue
        cx, cy, a, b = sd["ellipse"]
        px, py = view.back(cx, cy)
        ra, rb = (b, a) if view.rot in (90, 270) else (a, b)
        out += (f"q\n/Artifact <</Type /Layout /{TAG} ({sd['id']})>> BDC\n0 G {lw} w [] 0 d 0 J 0 j\n"
                f"{px + ra:.3f} {py:.3f} m\n"
                f"{px + ra:.3f} {py + K * rb:.3f} {px + K * ra:.3f} {py + rb:.3f} {px:.3f} {py + rb:.3f} c\n"
                f"{px - K * ra:.3f} {py + rb:.3f} {px - ra:.3f} {py + K * rb:.3f} {px - ra:.3f} {py:.3f} c\n"
                f"{px - ra:.3f} {py - K * rb:.3f} {px - K * ra:.3f} {py - rb:.3f} {px:.3f} {py - rb:.3f} c\n"
                f"{px + K * ra:.3f} {py - rb:.3f} {px + ra:.3f} {py - K * rb:.3f} {px + ra:.3f} {py:.3f} c\n"
                f"h S\nEMC\nQ\n").encode()
    return bytes(out)


PREFIX = b"q\n"


# ---------------------------------------------------------------------- 読み戻し
def drawn(doc, page_no: int) -> list[tuple[str, str, tuple[float, float, float, float], float]]:
    """書いた後のファイルのページに描いてある、とどくの○：(id, 形（"ellipse"・"border"）, 楕円 (cx, cy, a, b) か四角
    (x0, y0, x1, y1)（見た目の座標）, 線の太さ)。描画を読み直して、印の中の曲線の外枠と、線の太さ（w）から。"""
    pg = doc.pages()[page_no - 1]
    view = View(doc, pg)
    data = doc.page_content(pg)
    out = []
    for m in re.finditer(rb"/" + TAG.encode() + rb"\s*\(([^)]*)\)([^>]*)>>\s*BDC(.*?)EMC", data, re.S):
        pts, stack = [], []
        lw = 0.0
        for kind, val in tokens(m[3]):
            if kind == "n":
                stack.append(float(val))
                continue
            if kind == "o" and val == "w" and stack:
                lw = stack[-1]
            elif kind == "o" and val == "m" and len(stack) >= 2:
                pts.append((stack[-2], stack[-1]))
            elif kind == "o" and val == "c" and len(stack) >= 6:
                pts += [(stack[-6], stack[-5]), (stack[-4], stack[-3]), (stack[-2], stack[-1])]
            stack = []
        if len(pts) < 13:
            continue  # 4 つの曲線（楕円・四角）でない
        xs = [view.pt(x, y)[0] for x, y in pts]
        ys = [view.pt(x, y)[1] for x, y in pts]
        x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
        if b"/TdkStyle /border" in m[2]:
            out.append((m[1].decode("latin-1"), "border", (x0, y0, x1, y1), lw))
        else:
            out.append((m[1].decode("latin-1"), "ellipse", ((x0 + x1) / 2, (y0 + y1) / 2, (x1 - x0) / 2, (y1 - y0) / 2), lw))
    return out


def circled_text(lay, kind: str, g, lw: float = 0.0, loose: bool = False) -> tuple[list, list]:
    """ページの字のうち、囲みの形が囲んでいる字と、囲まずに（太さ lw の線が）触れている字（書くときと同じ決め方：encloses・hits）。
    書くときに ±TOL の丸めの分だけ縁の字に触れて見えないよう、線の太さは丸めの分（TOL）だけ細く見る。"""
    w = max(0.0, lw - 2 * TOL)
    # 囲んでいる字：線の内側の縁（太さ w）が字の箱をすっかり囲むもの（線が字の箱を切る○は、囲んだとしない）
    inner = [ch for ch in lay.chars if not ch.text.isspace() and encloses(kind, g, ch, w)]
    ids = {id(ch) for ch in inner}
    touch = [ch for ch in lay.chars if not ch.text.isspace() and id(ch) not in ids and hits(kind, g, ch, w, loose)]
    return inner, touch


def read_back(doc, circled: list[dict]) -> list[dict]:
    """書いた後のファイルで、○がそれぞれ、言った字を囲み、ほかの字に触れていないか（書くときと同じ決め方で）。違う所を返す。"""
    out = []
    layouts: dict = {}
    by_page: dict[int, dict] = {}
    for c in circled:
        n = c["page"]
        if n not in by_page:
            by_page[n] = {i: (k, g, w) for i, k, g, w in drawn(doc, n)}
        hit = by_page[n].get(c["id"])
        if hit is None:
            out.append({"id": c["id"], "written": f"○「{c['text']}」", "read_back": "書いた後のファイルのページに、この○が見つかりません"})
            continue
        kind, g, lw = hit
        lay = layout(doc, n, layouts)
        inner, touch = circled_text(lay, kind, g, lw, bool(c.get("loose")))
        got = "".join(ch.text for ch in sorted(inner, key=lambda ch: (ch.y0, ch.x0) if ch.vertical else (ch.x0, ch.y0)))
        want = re.sub(r"\s", "", c["text"])
        if c.get("unread"):
            # 字の読めない所に枠で描いた○：囲んだ字は比べない（結果で、そう伝えている）。ページの字に触れていないかだけ
            if touch or got:
                out.append({"id": c["id"], "written": f"○（枠で描いた所）", "read_back":
                            f"ページの字「{(got + ''.join(ch.text for ch in touch))[:12]}」にかかっています"})
            continue
        if re.sub(r"\s", "", got) != want or touch:
            what = f"囲んでいる字「{got}」" + (f"・触れている字「{''.join(ch.text for ch in touch[:6])}」" if touch else "")
            out.append({"id": c["id"], "written": f"○「{c['text']}」", "read_back": what})
    return out
