# -*- coding: utf-8 -*-
"""販売大将に取り込む売上データ（データ・インポート情報登録 ／ 売上伝票 パターン0001）を作る。

流れ: 一覧表の期間の行（GAS の action 'sales'）→ 画面で直せる表の行（edit_rows）→ 取り込み用CSV（to_csv）。
決まりは 2026-11-08 に、先方の取り込み設定の画面（スクショ）と売上エクスポート（9/1〜10/21）で確かめた。
  CSV・カンマ区切り・文字列の引用符なし・見出しなし・改行 CRLF・文字コード cp932・日付は yyyy/mm/dd・21列。
  商品コードは一覧表の商品CDの頭に「00」を付けたもの（9月の171件中170件で一致）。
  伝票は「売上日・得意先・仮納品書のタブ」ごと（同じ日・同じ得意先でも納品書ごとに分かれている）。
  ただし相手先NO（注文番号）や直送先が違う行は、伝票の番号が同じでも別の伝票にする（実物では注文ごとに伝票が分かれている）。
  部門は、西九条から出荷した仮納品書とバルク（単位 kg）を西九条工場、ほかを工場にする。
  2026-11-09 に部門コードの一覧を受け取り、9月の売上と突き合わせて確かめた（一覧表の行に当たる売上で、バルク以外137行は全部工場、
  バルク30行のうち27行が西九条工場）。事務所は有償サンプル・原料販売の手入力の行、取引先の工場（5000xx）は仮納品書の無い出荷で使われる。
"""
import re
import unicodedata

KUBUN = {"通常": "1", "無償": "8", "返品": "10", "記事": "18"}   # 取引区分＝取引性格
BUMON_HONSHA = "000002"     # 工場（本社の工場）
BUMON_NISHIKUJO = "000003"  # 西九条工場
BUMON_GUIDE = "000002 工場・000003 西九条工場・000001 事務所（有償サンプル・原料販売）・999999 未使用（返品など）・5000xx 取引先の工場"
COLUMNS = ["伝票", "売上日", "得意先コード", "得意先", "区分", "商品コード", "商品名", "数量", "単価", "原価単価",
           "相手先NO", "部門コード", "伝票備考", "明細備考", "商品摘要"]
DOC_FIELDS = ["売上日", "得意先コード", "部門コード", "伝票備考", "相手先NO"]   # 1枚の伝票の中で同じ。違えば伝票を分ける


def _norm(s):
    s = unicodedata.normalize("NFKC", str(s or ""))
    s = re.sub(r"株式会社|有限会社|合同会社|\(株\)|\(有\)|㈱|㈲|御中|様", "", s)   # NFKC で ㈱→(株)・㈲→(有)
    return re.sub(r"[\s・･.\-ー—―]", "", s).upper()


def _same_party(a, b):
    """送り先が得意先そのものか（法人格・敬称・記号の違いは無視。片方がもう片方に含まれれば同じ）"""
    x, y = _norm(a), _norm(b)
    return bool(x and y) and (x in y or y in x)


def _full_corp(s):
    """伝票備考用。送り先の ㈱・㈲ を「株式会社」「有限会社」に戻す（販売大将の伝票備考は「株式会社○○倉庫様直送」の形）"""
    s = unicodedata.normalize("NFKC", str(s or ""))
    return s.replace("(株)", "株式会社").replace("(有)", "有限会社")


def _num(v):
    if v is None or v == "":
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return int(f) if f == int(f) else f


def edit_rows(gas_rows):
    """一覧表の期間の行 → 画面で直せる表の行（COLUMNS の順の dict）"""
    out, docs = [], {}
    for r in gas_rows:
        key = (r.get("出荷日", ""), r.get("得意先コード", ""), r.get("仮納品書シート", ""))
        if key not in docs:
            docs[key] = len(docs) + 1
        qty, price = _num(r.get("出荷数")), _num(r.get("単価"))
        dest = str(r.get("送り先") or "").strip()
        cust = str(r.get("得意先") or "").strip()
        cd = str(r.get("商品CD") or "").strip()
        lot = str(r.get("ロット") or "").strip()
        if qty is not None and qty < 0:
            kubun = "返品"
        elif price == 0:
            kubun = "無償"
        else:
            kubun = "通常"
        out.append({
            "伝票": docs[key],
            "売上日": str(r.get("出荷日", "")).replace("-", "/"),
            "得意先コード": str(int(r["得意先コード"])) if str(r.get("得意先コード", "")).isdigit() else "",
            "得意先": cust,
            "区分": kubun,
            "商品コード": ("00" + cd) if cd else "",
            "商品名": str(r.get("製品名") or ""),
            "数量": qty,
            "単価": price,
            "原価単価": 0,
            "相手先NO": "",
            "部門コード": BUMON_NISHIKUJO if ("西九条" in str(r.get("ファイル") or "") or str(r.get("単位") or "").lower() == "kg") else BUMON_HONSHA,
            "伝票備考": ("%s様直送" % re.sub(r"(様|御中)$", "", _full_corp(dest))) if dest and not _same_party(dest, cust) else "",
            "明細備考": ("Lot.%s" % lot) if lot else "",
            "商品摘要": "",
        })
    return out


def _text(v):
    """CSV に入れる文字。引用符を使わない形式なので、カンマ・改行・引用符は置き換える"""
    s = "" if v is None else str(v)
    if s.lower() == "nan":
        return ""
    s = s.replace("\r", " ").replace("\n", " ").replace(",", "、").replace('"', "”")
    return s.strip()


def _cp932_ok(s):
    try:
        s.encode("cp932")
        return True
    except UnicodeEncodeError:
        return False


def problems(rows):
    """CSV にする前に直す必要がある点（[文]）。空なら書き出してよい"""
    out = []
    for i, r in enumerate(rows, 1):
        at = "%d行目（伝票%s）: " % (i, r.get("伝票", "?"))
        kubun = r.get("区分") or "通常"
        if kubun not in KUBUN:
            out.append(at + "区分「%s」は 通常・無償・返品・記事 のどれかにしてください" % kubun)
            continue
        if not re.fullmatch(r"\d{4}/\d{1,2}/\d{1,2}", str(r.get("売上日") or "")):
            out.append(at + "売上日を 2026/11/09 の形で入れてください")
        if not str(r.get("得意先コード") or "").strip().isdigit():
            out.append(at + "得意先コードを数字で入れてください")
        if not re.fullmatch(r"\d{6}", str(r.get("部門コード") or "").strip()):
            out.append(at + "部門コードを6桁の数字で入れてください（000002 工場 など）")
        if kubun != "記事":
            if not str(r.get("商品コード") or "").strip():
                out.append(at + "商品コードが空です（文字だけの行なら区分を「記事」に）")
            if _num(r.get("数量")) is None:
                out.append(at + "数量が空です")
            if kubun in ("通常", "返品") and _num(r.get("単価")) is None:
                out.append(at + "単価が空です（単価0のサンプルなら区分を「無償」に）")
        for f in ("商品摘要", "伝票備考", "明細備考", "相手先NO"):
            if not _cp932_ok(_text(r.get(f))):
                out.append(at + "「%s」に販売大将で使えない文字があります" % f)
    return out


def _fmt_num(v):
    n = _num(v)
    if n is None:
        return "0"
    return str(n) if isinstance(n, int) else ("%s" % round(n, 6)).rstrip("0").rstrip(".")


def to_csv(rows):
    """取り込み用CSV（cp932 の bytes）。伝票の番号順に並べ、伝票ごとに行番号を1から振る"""
    order = []
    docs = {}
    for r in rows:
        key = (r.get("伝票"),) + tuple(_text(r.get(f)) for f in DOC_FIELDS)   # 番号が同じでも項目が違えば別の伝票
        if key not in docs:
            docs[key] = []
            order.append(key)
        docs[key].append(r)
    lines = []
    for key in sorted(order, key=lambda k: (_num(k[0]) is None, _num(k[0]) or 0)):   # 同じ番号の中は出てきた順
        for gyo, r in enumerate(docs[key], 1):
            kubun = r.get("区分") or "通常"
            kb = KUBUN[kubun]
            kiji = kubun == "記事"
            lines.append(",".join([
                "0", "ur01", "0", _text(r.get("売上日")), _text(r.get("得意先コード")), str(gyo), kb, kb,
                "" if kiji else _text(r.get("商品コード")),
                "0" if kiji else "1", "0" if kiji else "1",
                "0" if kiji else _fmt_num(r.get("数量")),
                "0" if kiji else _fmt_num(r.get("単価")),
                "0" if kiji else _fmt_num(r.get("原価単価")),
                "0", "0" if kiji else "1",
                _text(r.get("相手先NO")), _text(r.get("部門コード")), _text(r.get("伝票備考")),
                _text(r.get("明細備考")), _text(r.get("商品摘要")),
            ]))
    return ("\r\n".join(lines) + "\r\n").encode("cp932", errors="replace") if lines else b""
