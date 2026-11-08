# -*- coding: utf-8 -*-
"""画面「販売大将に取り込む」（app.py から呼ばれる）
一覧表の期間の行を読み、販売大将の売上伝票インポート（パターン0001）のCSVにする。
先方の流れ（10/21 のメモ）: ①データを入れる → ②売上データを抽出 → ③開いて手修正（運賃・資材の追加、相手先NO、
ロット別単価での単価修正、部門コード）→ ④販売大将へインポート。③をこの画面の表で行う。
"""
import datetime
import importlib

import pandas as pd
import streamlit as st

import common
import uriage

importlib.reload(uriage)  # Streamlit Cloudの旧モジュール残留対策

common.apply_css()

st.title("📤 販売大将に取り込む")
st.caption("得意先ごと出荷一覧表の行から、販売大将に取り込む売上データ（CSV）を作ります。表で直してからダウンロードしてください。｜%s"
           % common.APP_VERSION)

with st.sidebar:
    st.header("使い方")
    st.markdown(
        "1. 売上を上げる **期間**（ふつうは出荷日の1日）を選んで **「一覧表から読み込む」**\n"
        "2. 表で直す：**単価**（ロット別単価で違うとき）・**相手先NO**・**部門コード**・**伝票備考**（直送先）\n"
        "3. 運賃・資材は、表の一番下に **行を足して** 入力（商品コード Z000000001 立替運賃 など。同じ伝票の番号にする）\n"
        "4. 赤い指摘が無くなったら **CSVをダウンロード** し、販売大将のデータ・インポート（売上伝票）で取り込む\n\n"
        "・伝票は「出荷日・得意先・仮納品書のタブ」ごとに分けています。まとめたいときは「伝票」の番号をそろえる"
        "（番号が同じでも、相手先NO・直送先・部門が違う行は自動で別の伝票になります）\n"
        "・部門コードは今は仮です（本社＝000002・西九条から出荷した分とバルク＝000003）\n"
        "・一覧表は変わりません"
    )

today = (datetime.datetime.utcnow() + datetime.timedelta(hours=9)).date()   # クラウドはUTCなので日本時間にする
c1, c2, c3 = st.columns([1, 1, 1])
start = c1.date_input("開始日（出荷日）", value=today, format="YYYY/MM/DD")
end = c2.date_input("終了日（出荷日）", value=today, format="YYYY/MM/DD")
c3.write("")
c3.write("")
load = c3.button("📥 一覧表から読み込む", type="primary")

if load:
    if start > end:
        st.error("開始日が終了日より後になっています。")
    else:
        with st.spinner("一覧表を読んでいます…"):
            res = common.post_ichiran({"action": "sales", "start": start.isoformat(), "end": end.isoformat()}, timeout=120)
        if not res.get("ok"):
            st.error("一覧表を読めませんでした：%s" % common.md_escape(res.get("error", "不明なエラー")))
        else:
            st.session_state["uriage"] = {"period": (start.isoformat(), end.isoformat()), "rows": uriage.edit_rows(res.get("rows", [])),
                                          "unassigned": res.get("unassigned", 0), "noDate": res.get("noDate", 0)}
            st.session_state.pop("uriage_editor", None)

data = st.session_state.get("uriage")
if data:
    s, e = data["period"]
    st.markdown("#### %s 〜 %s の売上（%d 行）" % (s.replace("-", "/"), e.replace("-", "/"), len(data["rows"])))
    if data["unassigned"]:
        st.warning("一覧表の「未振り分け」に、この期間の行が %d 行あります（ここには入っていません）。"
                   "一覧表で正しい得意先のシートへ移してから、もう一度読み込んでください。" % data["unassigned"])
    if not data["rows"]:
        st.info("この期間の行はありません。")
    else:
        st.caption("部門コードは仮です（本社＝000002・西九条から出荷した分とバルク＝000003）。部門コードの一覧が届いたら差し替えます。")
        edited = st.data_editor(
            pd.DataFrame(data["rows"], columns=uriage.COLUMNS), key="uriage_editor", hide_index=True,
            num_rows="dynamic", width="stretch",
            column_config={
                "伝票": st.column_config.NumberColumn("伝票", help="同じ番号の行が1枚の伝票になります", step=1, format="%d"),
                "区分": st.column_config.SelectboxColumn("区分", options=list(uriage.KUBUN), help="単価0のサンプルは無償・文字だけの行は記事"),
                "数量": st.column_config.NumberColumn("数量", format="%g"),
                "単価": st.column_config.NumberColumn("単価", format="%g"),
                "原価単価": st.column_config.NumberColumn("原価単価", format="%g", help="運賃・資材などで使う"),
                "得意先": st.column_config.TextColumn("得意先", help="確認用（CSVには入りません）"),
                "商品名": st.column_config.TextColumn("商品名", help="確認用（CSVには入りません）"),
                "伝票備考": st.column_config.TextColumn("伝票備考", help="直送先など"),
                "明細備考": st.column_config.TextColumn("明細備考", help="Lot.○○"),
                "商品摘要": st.column_config.TextColumn("商品摘要", help="新単価・新製品など"),
            })
        rows = [{k: (None if pd.isna(v) else v) for k, v in r.items()} for r in edited.to_dict("records")]
        probs = uriage.problems(rows)
        if probs:
            st.error("CSVにする前に直してください（%d 件）" % len(probs))
            for p in probs[:30]:
                st.markdown("- " + common.md_escape(p))
            if len(probs) > 30:
                st.caption("ほか %d 件" % (len(probs) - 30))
        name = "売上伝票インポート_%s〜%s.csv" % (s, e) if s != e else "売上伝票インポート_%s.csv" % s
        st.download_button("⬇️ 取り込み用CSVをダウンロード", data=uriage.to_csv(rows) if not probs else b"",
                           file_name=name, mime="text/csv", type="primary", disabled=bool(probs), on_click="ignore")
