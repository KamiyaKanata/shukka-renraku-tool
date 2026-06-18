# -*- coding: utf-8 -*-
"""出荷連絡表 自動生成 — Web版(MVP)
仮納品書 + 商品マスタ をアップロード → ボタン → 単価入り出荷連絡表をダウンロード。
起動: ./.venv/bin/streamlit run app.py
"""
import streamlit as st
import pandas as pd
import engine

st.set_page_config(page_title="出荷連絡表 自動生成", page_icon="📦", layout="wide")


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

st.title("📦 出荷連絡表 自動生成（MVP）")
st.caption("仮納品書と商品マスタをアップロードして「生成」を押すと、単価入りの出荷連絡表ができます。")

with st.sidebar:
    st.header("使い方")
    st.markdown(
        "1. **仮納品書**（Excel）をアップロード\n"
        "2. **商品マスタ**（商品単価リスト .xlsm / .xlsx / .csv）をアップロード\n"
        "3. **生成**ボタンを押す\n"
        "4. 表を確認して **ダウンロード**\n\n"
        "・資材／残資材は自動で除外、有償サンプルは含めます\n"
        "・単価は品名＋数量（ロット閾値）で自動ひも付け\n"
        "・引けない／曖昧な行は **🔶要確認** で色付け表示"
    )
    date_label = st.text_input("出荷日（空欄なら仮納品書から自動取得）", value="",
                               help="例: 2026-06-08。空欄のままなら仮納品書の日付を自動で使います。")

c1, c2 = st.columns(2)
with c1:
    kari = st.file_uploader("① 仮納品書（.xlsx / .xlsm）", type=["xlsx", "xlsm"])
with c2:
    master = st.file_uploader("② 商品マスタ（.xlsm / .xlsx / .csv）", type=["xlsm", "xlsx", "csv"])

go = st.button("🚀 出荷連絡表を生成", type="primary", disabled=not (kari and master))

if go and kari and master:
    try:
        with st.spinner("解析中…（マスタが大きい場合は数十秒かかります）"):
            rows, colorder, stats, debug = engine.generate(kari, master, master.name, date_label)
    except Exception as e:
        st.error(f"処理でエラーが発生しました: {e}")
        st.stop()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("出荷連絡表 行数", stats["出荷連絡表 行数"])
    m2.metric("要確認 行数", stats["要確認 行数"])
    m3.metric("仮納品書 明細(資材除外後)", stats["仮納品書 明細数(資材除外後)"])
    m4.metric("マスタ商品数", stats["マスタ商品数"])

    _src = "仮納品書から自動取得" if stats.get("日付自動取得") else ("手入力" if date_label.strip() else "本日")
    st.caption(f"📅 出荷日: **{stats['出荷日']}**（{_src}）")

    if stats["要確認 行数"]:
        st.warning(f"🔶 {stats['要確認 行数']} 行が「要確認」です（単価未取得・数量帯外・名前衝突など）。下表の色付き行をご確認ください。")
    else:
        st.success("✅ すべての行で単価をひも付けできました。")

    st.markdown("#### 出荷連絡表（プレビュー・全文表示）")
    df = pd.DataFrame(rows)[colorder]
    def _hl(row):
        return ['background-color:#FFE3EC' if row["要確認"] == "要確認" else '' for _ in row]
    # st.table は折り返して全行・全文を表示（途中で切れない）
    st.table(df.style.apply(_hl, axis=1).hide(axis="index"))

    bio = engine.to_workbook(rows, colorder, stats["出荷日"])
    st.download_button(
        "⬇️ 出荷連絡表をダウンロード",
        data=bio.getvalue(),
        file_name=f"出荷連絡表_{stats['出荷日']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    with st.expander(f"🔍 解析の内訳（読めた明細 {len(debug['items'])} 件 / 除外 {len(debug['excluded'])} 件）— 「商品が乗らない」原因の確認用"):
        st.markdown("**① 仮納品書から読み込んだ明細（資材除外後）** — ここに無い商品は仮納品書の読み取りで拾えていません")
        st.table(pd.DataFrame(debug["items"]) if debug["items"] else pd.DataFrame({"(なし)": []}))
        st.markdown("**② 除外した行（資材／残資材のみ）** — ※数量なしは除外せず、上の出荷連絡表に「要確認」で表示します")
        st.table(pd.DataFrame(debug["excluded"]) if debug["excluded"] else pd.DataFrame({"(なし)": []}))
else:
    st.info("①②の両方をアップロードすると「生成」できます。")
