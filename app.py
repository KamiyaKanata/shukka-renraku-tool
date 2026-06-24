# -*- coding: utf-8 -*-
"""出荷連絡表 自動生成 — Web版(MVP)
仮納品書 + 商品マスタ をアップロード → ボタン → 単価入り出荷連絡表をダウンロード。
起動: ./.venv/bin/streamlit run app.py
"""
import datetime
import importlib
import streamlit as st
import pandas as pd
import engine
importlib.reload(engine)  # Streamlit Cloudの旧モジュール残留対策（app.pyと常に同じ版に揃える）

st.set_page_config(page_title="出荷連絡表 自動生成", page_icon="📦", layout="wide")
_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _detail_preview_df(groups):
    """ダウンロードの『詳細シート』と同じ内容（得意先＝シート名／商品CD昇順／CD0補正）。"""
    dated = [(g["日付"], r) for g in groups for r in g["records"]]
    dated = sorted(dated, key=lambda dr: engine._cd_sortkey(dr[1]["商品CD"]))
    rows = [{
        "得意先": r.get("得意先", ""),
        "製品名": r["製品名"], "ロット": r["ロット"],
        "出荷数": ("" if r["数量"] is None else engine._num_out(r["数量"])),
        "ケース": engine._cases_str(r["cases"]),
        "商品CD": engine._pad_cd(r["商品CD"]), "処方番号": r["処方番号"],
        "単価": ("" if r["単価"] is None else r["単価"]),
        "要確認": ("要確認" if r["要確認"] else ""), "メモ": r["メモ"],
    } for dl, r in dated]
    return pd.DataFrame(rows) if rows else pd.DataFrame({"(なし)": []})


def _push_history(filename, data, summary):
    """このセッションの直近3件を保持（新しい順）。サーバーには保存しない。"""
    hist = st.session_state.setdefault("history", [])
    hist.insert(0, {"filename": filename, "data": data, "summary": summary})
    del hist[3:]


def require_password():
    """st.secrets['APP_PASSWORD'] が設定されていれば認証を要求。未設定(ローカル)なら素通り。"""
    try:
        expected = st.secrets["APP_PASSWORD"]
    except Exception:
        return True
    if st.session_state.get("auth_ok"):
        return True
    st.markdown("### 🔒 ログイン")
    pw = st.text_input("パスワード", type="password")
    if st.button("ログイン"):
        if pw == expected:
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("パスワードが違います")
    return False


if not require_password():
    st.stop()

APP_VERSION = "v3.0（印刷=CD順+曜日／詳細=得意先列・日付は見出し横）"
st.title("📦 出荷連絡表 自動生成（MVP）")
st.caption(f"仮納品書と商品マスタをアップロードして「生成」を押すと、単価入りの出荷連絡表ができます。｜{APP_VERSION}")

with st.sidebar:
    st.header("使い方")
    st.markdown(
        "1. **仮納品書**（Excel）をアップロード（**複数選択可**）\n"
        "2. **商品マスタ**（商品単価リスト .xlsm / .xlsx / .csv）をアップロード\n"
        "3. **生成**ボタンを押す\n"
        "4. 表を確認して **ダウンロード**\n\n"
        "・複数ファイル／複数シートは **出荷日ごとに別シート**で1つのExcelにまとめます\n"
        "・同一日付の中は製品×ロットで合算します\n"
        "・資材／残資材は自動で除外、有償サンプルは含めます\n"
        "・**赤文字（黒との二重記載の控え）は自動で除外**し二重計上しません\n"
        "・単価は品名＋数量（ロット閾値）で自動ひも付け\n"
        "・引けない／曖昧な行は **🔶要確認** として表示\n"
        "・直近3件は画面下の **履歴** から再ダウンロード可（このセッション内のみ）"
    )
    date_label = st.text_input("出荷日（空欄なら仮納品書から自動取得）", value="",
                               help="例: 2026-06-08。空欄のままなら仮納品書の日付を自動で使います。")

c1, c2 = st.columns(2)
with c1:
    kari = st.file_uploader("① 仮納品書（.xlsx / .xlsm）※複数選択可", type=["xlsx", "xlsm"],
                            accept_multiple_files=True)
with c2:
    master = st.file_uploader("② 商品マスタ（.xlsm / .xlsx / .csv）", type=["xlsm", "xlsx", "csv"])

go = st.button("🚀 出荷連絡表を生成", type="primary", disabled=not (kari and master))

if go and kari and master:
    try:
        with st.spinner("解析中…（マスタが大きい場合は数十秒かかります）"):
            groups, stats, debug = engine.generate(kari, master, master.name, date_label)
    except Exception as e:
        st.error(f"処理でエラーが発生しました: {e}")
        st.stop()

    dates = stats["日付一覧"]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("製品数", stats["出荷連絡表 行数"])
    m2.metric("要確認 製品数", stats["要確認 行数"])
    m3.metric("出荷日(シート)数", stats["シート数"])
    m4.metric("仮納品書ファイル数", stats["ファイル数"])

    _src = "仮納品書から自動取得" if stats.get("日付自動取得") else ("手入力" if date_label.strip() else "本日")
    st.caption(f"📅 出荷日: **{' / '.join(dates)}**（{_src}）／ マスタ商品数 {stats['マスタ商品数']}")
    if len(dates) > 1:
        st.info(f"🗂 複数の出荷日（{', '.join(dates)}）を検出 → **日付ごとに別シート**で1つのExcelにまとめました。")

    if stats["要確認 行数"]:
        st.warning(f"🔶 {stats['要確認 行数']} 行が「要確認」です（単価未取得・数量帯外・名前衝突など）。下表の該当行をご確認ください。")
    else:
        st.success("✅ すべての行で単価をひも付けできました。")

    # ダウンロード（日付ごとのシートを持つ1ブック）
    bio = engine.to_workbook(groups)
    fname = f"出荷連絡表_{dates[0]}.xlsx" if len(dates) == 1 else f"出荷連絡表_{dates[0]}〜{dates[-1]}.xlsx"
    st.download_button("⬇️ 出荷連絡表をダウンロード", data=bio.getvalue(), file_name=fname,
                       mime=_XLSX_MIME, type="primary")
    summary = f"{len(dates)}シート / {stats['出荷連絡表 行数']}行 / 要確認{stats['要確認 行数']}"
    _push_history(fname, bio.getvalue(), summary)

    # プレビュー＝ダウンロードの「詳細シート」と同じ内容。印刷用シートはダウンロードにのみ入る。
    st.markdown("#### 内容プレビュー（＝詳細シートと同じ。ダウンロードには別途『印刷用シート（実物フォーマット）』も入ります）")
    st.caption("出荷日: " + "、".join(f"{d}（{engine._weekday_jp(d)}）" for d in dates))
    st.table(_detail_preview_df(groups).style.hide(axis="index"))

    with st.expander(f"🔍 解析の内訳（読めた明細 {len(debug['items'])} 件 / 除外 {len(debug['excluded'])} 件）— 「商品が乗らない」原因の確認用"):
        if debug.get("per_file"):
            st.markdown("**ファイル別の読み取り結果**")
            st.table(pd.DataFrame(debug["per_file"]))
        st.markdown("**① 仮納品書から読み込んだ明細（資材除外後）** — ここに無い商品は仮納品書の読み取りで拾えていません")
        st.table(pd.DataFrame(debug["items"]) if debug["items"] else pd.DataFrame({"(なし)": []}))
        st.markdown("**② 除外した行（資材／残資材・赤文字の控え）** — ※数量なしは除外せず、上の出荷連絡表に「要確認」で表示します")
        st.table(pd.DataFrame(debug["excluded"]) if debug["excluded"] else pd.DataFrame({"(なし)": []}))
else:
    st.info("①②の両方をアップロードすると「生成」できます。")

# ===== 履歴（このセッションの直近3件・サーバー保存なし）=====
_hist = st.session_state.get("history", [])
if _hist:
    st.divider()
    st.markdown("### 🕘 履歴（このセッションの直近3件）")
    st.caption("※サーバーには保存されません。ブラウザを閉じる／アプリ再起動で消えます。")
    for i, h in enumerate(_hist):
        col1, col2 = st.columns([3, 1])
        col1.write(f"**{h['filename']}** — {h['summary']}")
        col2.download_button("⬇️ DL", data=h["data"], file_name=h["filename"],
                             mime=_XLSX_MIME, key=f"hist_{i}")
