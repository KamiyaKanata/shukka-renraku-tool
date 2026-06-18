# -*- coding: utf-8 -*-
"""出荷連絡表 自動生成エンジン
仮納品書(Excel) + 商品マスタ(xlsm/xlsx/csv) -> 単価入り出荷連絡表(xlsx)

ルール（要件定義 v1 / 2026-06-18 確定分）:
- 入力は仮納品書。タブ=納品先ごと。1製品が複数先に分納 -> 製品×ロットで数量合算。
- 資材(「残資材」ブロック・容器/ポンプ等)は除外。有償サンプルは含む。
- 1ファイル=その日の出荷連絡表1枚（日付フィルタなし）。受領書ブロックは二重計上を避けて無視。
- 単価は商品マスタから品名の正規化突合で取得（NFKC+空白除去+小文字）。
- 数量帯: ロット欄の下限数値を閾値とし「数量以下で最大の閾値」を採用。
         最小閾値未満なら最小単価を採用しつつ「要確認」。
- 価格履歴: 同条件で複数あれば最新の価格更新日を採用。
- 名前衝突(正規化後に別商品CDが複数): 「要確認」。
- 単価が引けない/曖昧は「要確認」で別掲（黙って誤単価を出さない）。
"""
import re
import csv
import io
import unicodedata
import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------------- 正規化ユーティリティ ----------------
def normalize(s):
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = re.sub(r"\s+", "", s)
    return s.lower()

def first_num(s):
    if s is None:
        return None
    s = unicodedata.normalize("NFKC", str(s))
    m = re.search(r"\d[\d,]*(?:\.\d+)?", s)
    return float(m.group(0).replace(",", "")) if m else None

def parse_date(s):
    if s is None:
        return None
    if isinstance(s, (datetime.date, datetime.datetime)):
        return s if isinstance(s, datetime.date) else s.date()
    s = unicodedata.normalize("NFKC", str(s))
    m = re.search(r"(\d{4})[./\-](\d{1,2})[./\-](\d{1,2})", s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None

# ---------------- 資材判定 ----------------
SHIZAI_EXACT = {"容器", "ポンプ", "キャップ", "化粧箱", "6コ箱", "外装", "ラベル",
                "シール", "中栓", "スパチュラ", "袋", "箱", "台紙", "パンフ",
                "リーフレット", "説明書", "個箱", "内箱", "外箱"}
_SHIZAI_NORM = {normalize(x) for x in SHIZAI_EXACT}

def is_shizai(name):
    if name is None:
        return True
    raw = str(name)
    if "残資材" in raw:
        return True
    return normalize(raw) in _SHIZAI_NORM

# ================= 商品マスタ読み込み =================
def _sheets_from(fileobj, filename):
    """ファイルを [ (sheet_name, rows) ] に変換。rows は list[tuple(cells)]。"""
    name = (filename or "").lower()
    if name.endswith(".csv"):
        data = fileobj.read()
        if isinstance(data, bytes):
            for enc in ("utf-8-sig", "utf-8", "cp932"):
                try:
                    text = data.decode(enc); break
                except UnicodeDecodeError:
                    text = data.decode("utf-8", "replace")
        else:
            text = data
        rows = list(csv.reader(io.StringIO(text)))
        return [("csv", rows)]
    wb = load_workbook(fileobj, data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        out.append((ws.title, list(ws.iter_rows(values_only=True))))
    return out

def _find_master_header(rows):
    """製品名・単価列を含むヘッダー行を探し、列マップを返す。"""
    for i, row in enumerate(rows[:30]):
        nc = [normalize(c) for c in row]
        has_name = any(("製品名" in c) or ("販売名" in c) for c in nc)
        has_price = any("単価" in c for c in nc)
        if has_name and has_price:
            colmap = {}
            for j, c in enumerate(nc):
                if (("製品名" in c) or ("販売名" in c)) and "製品名" not in colmap:
                    colmap["製品名"] = j
                if (("商品cd" in c) or ("商品コード" in c)) and "商品CD" not in colmap:
                    colmap["商品CD"] = j
                if "試作番号" in c and "試作番号" not in colmap:
                    colmap["試作番号"] = j
                if ("売上単価" in c or c == "単価") and "単価" not in colmap:
                    colmap["単価"] = j
                if "価格更新" in c and "価格更新" not in colmap:
                    colmap["価格更新"] = j
                if c == "ロット" and "ロット" not in colmap:
                    colmap["ロット"] = j
            if "製品名" in colmap and "単価" in colmap:
                return i, colmap
    return None, None

def load_master(fileobj, filename):
    """商品マスタを読み込み、正規化製品名 -> エントリ一覧 の索引を返す。"""
    index = {}
    n_entries = 0
    for sheet_name, rows in _sheets_from(fileobj, filename):
        hidx, colmap = _find_master_header(rows)
        if hidx is None:
            continue
        cn, cc = colmap["製品名"], colmap.get("商品CD")
        cs, cp = colmap.get("試作番号"), colmap["単価"]
        cu, cl = colmap.get("価格更新"), colmap.get("ロット")
        for row in rows[hidx + 1:]:
            if cn >= len(row):
                continue
            name = row[cn]
            if name is None or not str(name).strip():
                continue
            tanka = first_num(row[cp]) if cp < len(row) else None
            if tanka is None:
                continue
            entry = {
                "製品名": str(name).strip(),
                "商品CD": str(row[cc]).strip() if (cc is not None and cc < len(row) and row[cc] is not None) else "",
                "試作番号": str(row[cs]).strip() if (cs is not None and cs < len(row) and row[cs] is not None) else "",
                "単価": tanka,
                "価格更新": parse_date(row[cu]) if (cu is not None and cu < len(row)) else None,
                "ロット下限": first_num(row[cl]) if (cl is not None and cl < len(row)) else None,
                "シート": sheet_name,
            }
            index.setdefault(normalize(name), []).append(entry)
            n_entries += 1
    return index, n_entries

# ================= 単価選択 =================
def select_price(entries, qty):
    flags = []
    es = [e for e in entries if e["単価"] is not None]
    if not es:
        return None, ["単価が空"]
    thresholds = sorted({e["ロット下限"] for e in es if e["ロット下限"] is not None})
    if thresholds:
        le = [t for t in thresholds if t <= qty]
        if le:
            chosen_t = max(le)
        else:
            chosen_t = min(thresholds)
            flags.append("数量%d が最小ロット%d 未満→最小単価採用(要確認)" % (int(qty), int(chosen_t)))
        cands = [e for e in es if e["ロット下限"] == chosen_t]
    else:
        cands = es
    cds = {e["商品CD"] for e in cands if e["商品CD"]}
    if len(cds) > 1:
        flags.append("商品CD複数 %s →名前衝突(要確認)" % sorted(cds))
    chosen = max(cands, key=lambda e: e["価格更新"] or datetime.date(1900, 1, 1))
    return chosen, flags

# ================= 仮納品書 読み込み =================
def parse_karinouhin(fileobj):
    wb = load_workbook(fileobj, data_only=True, read_only=True)
    items = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        hidx, cmap = None, None
        for i, row in enumerate(rows):
            nc = [normalize(c) for c in row]
            if any("品名" in c for c in nc) and any("数量" in c for c in nc):
                cmap = {}
                for j, c in enumerate(nc):
                    if "品名" in c and "品名" not in cmap:
                        cmap["品名"] = j
                    if "数量" in c and "数量" not in cmap:
                        cmap["数量"] = j
                    if ("lot" in c or "ロット" in c) and "lot" not in cmap:
                        cmap["lot"] = j
                    if "備考" in c and "備考" not in cmap:
                        cmap["備考"] = j
                hidx = i
                break
        if hidx is None or "品名" not in cmap or "数量" not in cmap:
            continue
        pcol, qcol = cmap["品名"], cmap["数量"]
        lcol, bcol = cmap.get("lot"), cmap.get("備考")
        for row in rows[hidx + 1:]:
            name = row[pcol] if pcol < len(row) else None
            nn = normalize(name)
            if "合計" in nn or "受領" in nn:
                break  # 仮納品書ブロックの終端（受領書は無視）
            if not nn or is_shizai(name):
                continue
            qty = first_num(row[qcol]) if qcol < len(row) else None
            if qty is None:
                continue
            lot = row[lcol] if (lcol is not None and lcol < len(row)) else None
            bikou = row[bcol] if (bcol is not None and bcol < len(row)) else None
            items.append({
                "製品名": str(name).strip(),
                "数量": qty,
                "Lot": str(lot).strip() if lot else "",
                "備考": str(bikou).strip() if bikou else "",
                "シート": ws.title,
            })
    return items

# ================= 集計・突合 =================
def aggregate(items):
    agg = {}
    for it in items:
        key = (normalize(it["製品名"]), it["Lot"])
        if key not in agg:
            agg[key] = {"製品名": it["製品名"], "Lot": it["Lot"], "数量": 0.0, "備考": set()}
        agg[key]["数量"] += it["数量"]
        if it["備考"]:
            agg[key]["備考"].add(it["備考"])
    return list(agg.values())

def build_rows(agg, master_index):
    out = []
    for a in sorted(agg, key=lambda x: normalize(x["製品名"])):
        entries = master_index.get(normalize(a["製品名"]), [])
        chosen, flags = (None, ["単価リストに該当なし"]) if not entries else select_price(entries, a["数量"])
        tanka = chosen["単価"] if chosen else None
        out.append({
            "製品名": a["製品名"],
            "ロット": a["Lot"],
            "出荷数": int(a["数量"]),
            "ケース構成": " / ".join(sorted(a["備考"])),
            "商品CD": chosen["商品CD"] if chosen else "",
            "処方番号": chosen["試作番号"] if chosen else "",
            "単価": int(tanka) if tanka else None,
            "金額": int(tanka * a["数量"]) if tanka else None,
            "要確認": "要確認" if flags else "",
            "メモ": " / ".join(flags),
        })
    return out

# ================= 出力(xlsx) =================
HDR = ["製品名", "ロット", "出荷数", "ケース構成", "商品CD", "処方番号", "単価", "金額", "要確認", "メモ"]

def to_workbook(rows, date_label=""):
    wb = Workbook()
    ws = wb.active
    ws.title = "出荷連絡表(単価入り)"
    brand = PatternFill("solid", fgColor="B83280")
    hdrfill = PatternFill("solid", fgColor="AD1457")
    warn = PatternFill("solid", fgColor="FFEBEE")
    thin = Side(style="thin", color="E7C9D6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.merge_cells("A1:J1")
    ws["A1"] = "出荷連絡表（単価入り）　%s" % date_label
    ws["A1"].fill = brand
    ws["A1"].font = Font(name="Yu Gothic", color="FFFFFF", bold=True, size=14)
    ws["A1"].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 30
    for j, h in enumerate(HDR, 1):
        c = ws.cell(row=2, column=j, value=h)
        c.fill = hdrfill; c.font = Font(name="Yu Gothic", color="FFFFFF", bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center"); c.border = border
    for i, r in enumerate(rows, 3):
        for j, k in enumerate(HDR, 1):
            c = ws.cell(row=i, column=j, value=r[k])
            c.font = Font(name="Yu Gothic", size=10.5); c.border = border
            if k in ("単価", "金額", "出荷数"):
                c.number_format = "#,##0"; c.alignment = Alignment(horizontal="right")
        if r["要確認"]:
            for j in range(1, len(HDR) + 1):
                ws.cell(row=i, column=j).fill = warn
    widths = [34, 10, 9, 18, 10, 14, 9, 12, 8, 40]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(row=2, column=j).column_letter].width = w
    ws.freeze_panes = "A3"
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return bio

# ================= 一括処理 =================
def generate(karinouhin_fileobj, master_fileobj, master_filename, date_label=""):
    master_index, n = load_master(master_fileobj, master_filename)
    items = parse_karinouhin(karinouhin_fileobj)
    agg = aggregate(items)
    rows = build_rows(agg, master_index)
    stats = {
        "マスタ商品数": n,
        "仮納品書 明細数(資材除外後)": len(items),
        "出荷連絡表 行数": len(rows),
        "要確認 行数": sum(1 for r in rows if r["要確認"]),
    }
    return rows, stats

if __name__ == "__main__":
    import sys, json
    kari, master = sys.argv[1], sys.argv[2]
    out = sys.argv[3] if len(sys.argv) > 3 else "出力_出荷連絡表.xlsx"
    with open(master, "rb") as mf:
        midx, n = load_master(mf, master)
    with open(kari, "rb") as kf:
        items = parse_karinouhin(kf)
    rows = build_rows(aggregate(items), midx)
    print("マスタ商品数:", n, "/ 仮納品書明細:", len(items), "/ 出力行:", len(rows))
    print("-" * 90)
    for r in rows:
        print("%-30s Lot:%-6s 数:%6d 単価:%s 金額:%s %s" % (
            r["製品名"][:30], r["ロット"], r["出荷数"],
            str(r["単価"]), str(r["金額"]), ("[" + r["メモ"] + "]") if r["要確認"] else ""))
    bio = to_workbook(rows)
    with open(out, "wb") as f:
        f.write(bio.read())
    print("-" * 90); print("saved:", out)
