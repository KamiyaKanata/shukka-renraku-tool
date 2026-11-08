# -*- coding: utf-8 -*-
"""送り状PDF（仮納品書＋送り状を1枚にスキャンしたもの）から届先を読む。

1ページの作り: 上半分＝仮納品書（横書き）、下半分＝西濃の送り状（90度回っている。向きはページで違う）。
読み取り（OCR）は差し替えられるようにし、ここには「読めた文字から届先を決めるルール」だけを置く。
ocr は画像をまとめて受け取り、画像ごとの読めた文字を返す関数（本番は一覧表のGAS経由の Google ドライブの文字読み取り）。
AI は使わない。読み間違いは得意先名などの候補に寄せ、寄せられなければ人の確認に回す。
"""
import difflib
import re
import unicodedata

LABEL_TOP = 0.54        # ページの高さのうち、ここから下が送り状（10/5・10/6 の実物で確認）
SNAP_MIN = 0.72         # 候補に寄せるときの一致率の下限
SAME_MIN = 0.80         # 宛先と届先を「同じ」と見る一致率の下限

_ANCHORS = re.compile(r"送り状|受領票|元払|一般|TEL|出荷日|配達|取扱")
_NOT_NAME = re.compile(
    r"送り状|受領票|元払|一般|TEL|TE1|出荷日|配達|取扱|注意|指定|お届|届け|依頼|荷送|問合|問い合|受付|受領|西濃|運輸|"
    r"個口|記事|品名|重量|サンブルーム|午前|午後|月\d|\d月|年\d|\d年|〒|万円|保険|運賃|^化粧品$|^資材$|^様$"
)
_ADDRESS = re.compile(r"(都|道|府|県).*(市|区|町|村|郡)|\d+[-ー－−]\d+")
_CORP = re.compile(r"株式会社|有限会社|合同会社|合資会社|\(株\)|\(有\)|㈱|㈲")


def norm(s):
    """比べる用: 全角半角・空白・記号・法人格・敬称をそろえる"""
    s = unicodedata.normalize("NFKC", str(s or ""))
    s = _CORP.sub("", s)
    s = re.sub(r"[\s・･.\-ー—―_|\"'`\[\]{}#¢]", "", s)
    s = re.sub(r"(御中|様)$", "", s)
    return s.upper()


def ratio(a, b):
    """2つの名前の近さ（0〜1）。片方がもう片方に含まれていれば、長さの差を少しだけ割り引く"""
    x, y = norm(a), norm(b)
    if not x or not y:
        return 0.0
    if x == y:
        return 1.0
    r = difflib.SequenceMatcher(None, x, y).ratio()
    short, long_ = sorted((x, y), key=len)
    if len(short) >= 3 and short in long_:
        r = max(r, 0.85)
    return r


def clean_line(t):
    """OCRの行の前後に付くゴミ（罫線の読み間違いなど）を外す"""
    t = unicodedata.normalize("NFKC", str(t or "")).replace(" ", "")
    t = re.sub(r"^[^\w(（]+", "", t)
    t = re.sub(r"[^\w)）]+$", "", t)
    return re.sub(r"様$", "", t)


def label_score(lines):
    """送り状として正しい向きで読めているほど大きくなる（向きの判定用）"""
    return sum(1 for r in lines if _ANCHORS.search(unicodedata.normalize("NFKC", r["t"])))


def _is_name_like(t):
    if len(norm(t)) < 2 or _NOT_NAME.search(t) or _ADDRESS.search(t):
        return False
    digits = sum(c.isdigit() for c in t)
    return digits <= len(t) * 0.3


def snap(name, candidates):
    """読み取った名前を候補（その日の宛先・得意先名・過去の直送先）に寄せる。返り値: (候補, 一致率)。寄せられなければ (name, 0)"""
    best, best_r = name, 0.0
    for c in candidates:
        r = ratio(name, c)
        if r > best_r:
            best, best_r = c, r
    return (best, best_r) if best_r >= SNAP_MIN else (name, 0.0)


def judge(atena, todoke):
    """仮納品書の宛先と送り状の届先を比べる。返り値: '同じ' / '違う' / '読めない'"""
    if not todoke:
        return "読めない"
    return "同じ" if ratio(atena, todoke) >= SAME_MIN else "違う"


def match_tab(lines, tabs):
    """仮納品書（上半分）のOCR結果を、Excel のタブと結び付ける。宛先と、品名・ロットが読めた割合で点を付ける。
    同じ宛先のタブが2つあっても品名で分かれる。tabs: [{"タブ", "宛先", "品": [{"製品名", "Lot"}]}]
    返り値: (タブ, 点)。どのタブにも合わなければ (None, 点)"""
    text = norm("".join(r["t"] for r in lines))
    def hit(word):
        w = norm(word)
        if len(w) < 2:
            return False
        if w in text:
            return True
        return any(ratio(word, r["t"]) >= SNAP_MIN for r in lines)
    best, best_score = None, 0.0
    for tab in tabs:
        items = tab.get("品") or []
        found = sum(1 for it in items if hit(it.get("製品名", "")) or hit(it.get("Lot", "")))
        score = (1.0 if hit(tab.get("宛先", "")) else 0.0) + (found / len(items) if items else 0.0)
        if score > best_score:
            best, best_score = tab, score
    return (best, best_score) if best_score >= 1.0 else (None, best_score)


# ===================== Google ドライブの文字読み取り（文字だけで行の大きさは取れない）向け =====================
# 2026-11-08 の試験: 届先は「〒の行の次の行」と「お届け先」欄に出る。向きが合わないと長い片仮名が化けるので2向きを読む。
_POSTAL = re.compile(r"(〒|T)\s?\d{3}-?\d{4}")
MAX_PAGES = 40


def text_lines(text):
    """OCRの文字を行のリストに（match_tab・label_score に渡せる形）"""
    return [{"t": l.strip(), "h": 0} for l in str(text or "").splitlines() if l.strip()]


def todoke_lines(text):
    """送り状のOCRの文字から、届先らしい行を集める（〒の行の次にある名前らしい行／「お届け先」欄の文字）"""
    lines = [l["t"] for l in text_lines(text)]
    out = []
    for i, l in enumerate(lines):
        n = unicodedata.normalize("NFKC", l)
        if _POSTAL.search(n):   # 次の3行のうち名前らしい行を2つまで（読み違いのゴミが先に来ることがある）
            got = 0
            for j in range(i + 1, min(i + 4, len(lines))):
                t = clean_line(lines[j])
                if _is_name_like(t):
                    out.append(t)
                    got += 1
                    if got == 2:
                        break
        m = re.match(r"^お届け?先(.+)$", n.replace(" ", ""))
        if m and len(norm(m.group(1))) >= 2:
            out.append(clean_line(m.group(1)))
    return out


def pick_todoke_text(text, candidates):
    """文字だけの読み取り結果から届先を決める。既知の名前に一番よく寄る行を使う。
    返り値: (名前, 一致率。既知の名前でなければ 0)"""
    raws = todoke_lines(text)
    best, best_r = "", 0.0
    for r in raws:
        name, ratio_ = snap(r, candidates)
        if ratio_ > best_r:
            best, best_r = name, ratio_
    if best_r:
        return best, best_r
    return (raws[0] if raws else ""), 0.0


def page_images(pdf_bytes, dpi=300):
    """PDFの各ページから、上半分（仮納品書）と、送り状を90度・270度に回した白黒PNG（base64）を作る"""
    import base64
    import io
    import fitz
    from PIL import Image

    def png(im):
        buf = io.BytesIO()
        im.save(buf, "PNG", optimize=True)
        return base64.b64encode(buf.getvalue()).decode("ascii")

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    if doc.page_count > MAX_PAGES:
        raise ValueError("送り状PDFは %d ページまでにしてください（%d ページあります）" % (MAX_PAGES, doc.page_count))
    pages = []
    for i, page in enumerate(doc):
        W, H = page.rect.width, page.rect.height

        def crop(rect):
            pm = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), clip=rect, colorspace=fitz.csGRAY)
            im = Image.frombytes("L", (pm.width, pm.height), pm.samples)
            return im.point(lambda v: 255 if v > 128 else 0).convert("1")

        top = crop(fitz.Rect(0, 0, W, H * LABEL_TOP))
        label = crop(fitz.Rect(0, H * LABEL_TOP, W, H))
        pages.append({"page": i + 1, "top": png(top),
                      "l90": png(label.rotate(-90, expand=True)), "l270": png(label.rotate(-270, expand=True))})
    return pages


def tabs_from_items(items):
    """仮納品書の明細（engine.parse_karinouhin の結果）から、タブごとの宛先と品名・ロットをまとめる"""
    tabs = {}
    for it in items:
        t = tabs.setdefault(it["シート"], {"タブ": it["シート"], "宛先": it.get("発送先", ""), "品": []})
        t["品"].append({"製品名": it["製品名"], "Lot": it.get("Lot", "")})
    return list(tabs.values())


def read_okurijo(pdf_bytes, tabs, ocr, extra_candidates=(), progress=None):
    """送り状PDFを読む。ocr(images, want_candidates) は {"texts": {key: 文字}, "candidates": [...]} を返す関数。
    返り値: ページごとの行 [{"ページ", "タブ", "仮納品書の宛先", "送り状の届先", "一致率", "判定"}]"""
    pages = page_images(pdf_bytes)
    cands = None
    rows = []
    for k, p in enumerate(pages):
        if progress:
            progress(k, len(pages))
        n = p["page"]
        imgs = [{"key": "p%d_%s" % (n, x), "data": p[x]} for x in ("top", "l90", "l270")]
        res = ocr(imgs, cands is None)
        texts = res.get("texts", {})
        if cands is None:
            seen, cands = set(), []
            for c in [t["宛先"] for t in tabs] + list(res.get("candidates") or []) + list(extra_candidates):
                c = str(c or "").strip()
                if c and c not in seen:
                    seen.add(c)
                    cands.append(c)   # その日の宛先を先に置く（同じ一致率ならこちらを選ぶ）
        tab, _ = match_tab(text_lines(texts.get("p%d_top" % n, "")), tabs)
        best = None   # 2向きのうち、既知の名前に一番よく寄る方（同じなら送り状らしく読めた方）
        for x in ("l90", "l270"):
            t = texts.get("p%d_%s" % (n, x), "")
            name, r = pick_todoke_text(t, cands)
            key = (r, label_score(text_lines(t)))
            if name and (best is None or key > best[0]):
                best = (key, name, r)
        name, r = (best[1], best[2]) if best else ("", 0.0)
        atena = tab["宛先"] if tab else ""
        rows.append({"ページ": n, "タブ": tab["タブ"] if tab else "", "仮納品書の宛先": atena,
                     "送り状の届先": name, "一致率": round(r, 2),
                     "判定": ("タブ不明" if not tab else judge(atena, name))})
    return rows


def default_dest(row):
    """確認の表に最初から入れておく送り先。同じ・読めない → 仮納品書の宛先／違う → 送り状の届先"""
    if row.get("判定") == "違う" and row.get("送り状の届先"):
        return row["送り状の届先"]
    return row.get("仮納品書の宛先", "")


def decide(rows):
    """確認の表（画面で直したあとの行）から、タブごとの送り先とメモを決める。返り値: {タブ: {"送り先", "メモ"}}。
    仮納品書の宛先と同じなら engine 側で何もしない（メモも付かない）。
    読み取った名前が既知の名前でなく（一致率0）、そのまま使ったときは「確かめてください」を付ける。"""
    by_tab = {}
    for r in rows:
        tab = str(r.get("タブ") or "").strip()
        name = str(r.get("送り先") or "").strip()
        if tab and name:
            by_tab.setdefault(tab, []).append((name, r))
    out = {}
    for tab, items in by_tab.items():
        names = []
        for name, _ in items:
            if name not in names:
                names.append(name)
        atena = items[0][1].get("仮納品書の宛先", "")
        memo = ("仮納品書の宛先：%s（送り状の届先を送り先にした）" % atena) if atena else "送り状の届先を送り先にした"
        if any(not r.get("一致率") and name == r.get("送り状の届先") and r.get("判定") == "違う" for name, r in items):
            memo += " / 送り状の読み取りを確かめてください"
        out[tab] = {"送り先": "／".join(names), "メモ": memo}
    return out
