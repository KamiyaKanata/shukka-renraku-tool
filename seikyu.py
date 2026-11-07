# -*- coding: utf-8 -*-
"""販売大将の請求書CSVを読み、得意先ごと出荷一覧表のGASへ送る形（請求書ごとの明細）にする。

照合そのもの（得意先×商品コードで数量を比べる）は一覧表のGAS側で行う。ここでは読むだけ。
- 文字コードは UTF-8 で読めなければ cp932（販売大将の既定）
- 列は見出しの名前で探す（列の位置を決め打ちしない）
- 1行＝請求書の1明細。入金の行（売上の項目が空）は外して数える
"""
import csv
import datetime
import io
import re
import unicodedata

# 照合に必ず要る列（見出しの名前）
REQUIRED = ["請求書NO", "得意先コード", "得意先名称", "請求期間（自）", "請求期間（至）",
            "売上伝票日付", "売上伝票番号", "売上取引区分名", "商品コード", "商品名", "数量"]
OPTIONAL = ["単位名", "売上単価", "売上金額", "売上明細備考", "商品摘要"]


def _norm(s):
    return re.sub(r"\s", "", unicodedata.normalize("NFKC", str(s or "")))


def _decode(data):
    for enc in ("utf-8-sig", "cp932"):
        try:
            return data.decode(enc), enc
        except UnicodeDecodeError:
            continue
    raise ValueError("文字コードを読めません（UTF-8 でも Shift_JIS でもありません）")


def _iso(s):
    t = unicodedata.normalize("NFKC", str(s or "")).strip()
    if not t:
        return ""
    m = re.match(r"^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})", t)
    if not m:
        return None
    try:
        return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3))).isoformat()
    except ValueError:
        return None


def _num(s):
    t = unicodedata.normalize("NFKC", str(s or "")).strip().replace(",", "")
    if not t:
        return None
    try:
        v = float(t)
    except ValueError:
        return "NG"
    return int(v) if v.is_integer() else v


def parse_csv(data, filename=""):
    """請求書CSV（bytes）→ {"invoices": [...], "stats": {...}, "problem": ""}。
    problem が空でなければ照合に使えない（理由の文）。"""
    try:
        text, enc = _decode(data)
    except ValueError as e:
        return {"invoices": [], "stats": {}, "problem": "%s: %s" % (filename, e)}
    rows = list(csv.reader(io.StringIO(text)))
    if not rows:
        return {"invoices": [], "stats": {}, "problem": "%s: 中身が空です" % filename}
    head = [_norm(h) for h in rows[0]]
    col = {}
    for name in REQUIRED + OPTIONAL:
        key = _norm(name)
        if key in head:
            col[name] = head.index(key)
    missing = [n for n in REQUIRED if n not in col]
    if missing:
        return {"invoices": [], "stats": {}, "problem": "%s: 請求書CSVの見出しに %s がありません（販売大将の請求書の印刷画面から CSV で出したものか確かめてください）"
                % (filename, "・".join(missing))}

    def get(r, name):
        i = col.get(name)
        return r[i].strip() if i is not None and i < len(r) else ""

    invoices, order = {}, []
    skipped = {}
    bad = []
    n_lines = 0
    for no, r in enumerate(rows[1:], start=2):
        if not any(c.strip() for c in r):
            continue
        kubun = get(r, "売上取引区分名")
        code = get(r, "商品コード")
        if not kubun and not code:   # 入金の行など（売上の項目が空）
            skipped["入金など（売上の項目が空）"] = skipped.get("入金など（売上の項目が空）", 0) + 1
            continue
        start, end = _iso(get(r, "請求期間（自）")), _iso(get(r, "請求期間（至）"))
        date = _iso(get(r, "売上伝票日付"))
        qty = _num(get(r, "数量"))
        price, amount = _num(get(r, "売上単価")), _num(get(r, "売上金額"))
        if not start or not end or date is None or "NG" in (qty, price, amount):
            bad.append(no)
            continue
        key = (get(r, "請求書NO"), get(r, "得意先コード"), start, end)
        if key not in invoices:
            invoices[key] = {"請求書NO": key[0], "得意先コード": key[1], "得意先名称": get(r, "得意先名称"),
                             "自": start, "至": end, "lines": []}
            order.append(key)
        invoices[key]["lines"].append({
            "伝票日付": date, "伝票番号": get(r, "売上伝票番号"), "区分": kubun,
            "商品コード": code, "商品名": get(r, "商品名"), "摘要": get(r, "商品摘要"),
            "備考": get(r, "売上明細備考"), "数量": qty, "単位": get(r, "単位名"),
            "単価": price, "金額": amount,
        })
        n_lines += 1
    if bad:
        return {"invoices": [], "stats": {}, "problem": "%s: %d 行目などの日付・数値が読めません（%d 行）"
                % (filename, bad[0], len(bad))}
    inv = [invoices[k] for k in order]
    stats = {
        "ファイル": filename, "文字コード": enc, "請求書": len(inv), "明細": n_lines,
        "得意先": len({v["得意先コード"] for v in inv}),
        "請求期間": sorted({"%s〜%s" % (v["自"], v["至"]) for v in inv}),
        "外した行": skipped,
    }
    return {"invoices": inv, "stats": stats, "problem": ""}


def merge(results):
    """複数のCSVの結果をまとめる。同じ請求書（請求書NO・得意先・期間）が2つのファイルにあれば後の方を外す。"""
    seen, out, dup = set(), [], 0
    for res in results:
        for v in res["invoices"]:
            k = (v["請求書NO"], v["得意先コード"], v["自"], v["至"])
            if k in seen:
                dup += 1
                continue
            seen.add(k)
            out.append(v)
    return out, dup
