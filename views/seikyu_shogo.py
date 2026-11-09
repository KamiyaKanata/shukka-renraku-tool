# -*- coding: utf-8 -*-
"""画面「請求書と照合」（app.py から呼ばれる）
販売大将の請求書CSVを入れる → 得意先ごと出荷一覧表のGASが、得意先×商品コードの数量の合計で比べ、
結果を一覧表と同じフォルダの新しいスプレッドシートに書く。この画面には件数と得意先ごとの判定を出す。
"""
import importlib

import pandas as pd
import streamlit as st

import common
import seikyu

importlib.reload(seikyu)  # Streamlit Cloudの旧モジュール残留対策

common.apply_css()

st.title("🧾 請求書と照合")
st.caption("販売大将の請求書CSVを入れると、得意先ごと出荷一覧表と数量を比べて、ズレを一覧にします。｜%s" % common.APP_VERSION)

with st.sidebar:
    st.header("使い方")
    st.markdown(
        "1. 販売大将で請求書を出し、印刷画面から **CSV** で保存\n"
        "2. ここに **請求書CSV** をアップロード（20日締めと末締めなど、複数可）\n"
        "3. **「一覧表と照合する」** を押す\n"
        "4. 結果のファイルで、色の付いた行を確かめる\n\n"
        "・比べるのは **得意先ごと・商品コードごとの数量の合計** です（単価は見ません）\n"
        "・ロット違いで分かれた行は合計して比べます\n"
        "・運賃・資材と、請求しない出荷（ご支給原料の返却・原料のサンプル・外注先への送付）は **照合対象外** として結果に残します\n"
        "・一覧表は変わりません。結果は一覧表と同じフォルダに新しいファイルでできます"
    )

files = st.file_uploader("請求書CSV（販売大将の請求書の印刷画面から CSV で出したもの）※複数選択可",
                         type=["csv"], accept_multiple_files=True)

invoices, problems = [], []
if files:
    results = [seikyu.parse_csv(f.getvalue(), f.name) for f in files]
    problems = [r["problem"] for r in results if r["problem"]]
    for p in problems:
        st.error(common.md_escape(p))
    good = [r for r in results if not r["problem"]]
    if good:
        invoices, dup = seikyu.merge(good)
        st.table(pd.DataFrame([{
            "ファイル": r["stats"]["ファイル"], "請求書": r["stats"]["請求書"], "明細": r["stats"]["明細"],
            "得意先": r["stats"]["得意先"], "請求期間": "、".join(r["stats"]["請求期間"]),
            "外した行（入金など）": sum(r["stats"]["外した行"].values()),
        } for r in good]).style.hide(axis="index"))
        if dup:
            st.info("同じ請求書が複数のファイルに入っていたため、%d 件は1つにまとめました。" % dup)

go = st.button("🔎 一覧表と照合する", type="primary", disabled=not invoices or bool(problems))
if go and invoices:
    with st.spinner("照合中です（数十秒かかることがあります）…"):
        res = common.post_ichiran({"action": "reconcile", "invoices": invoices,
                                   "files": [f.name for f in files]}, timeout=300)
    if res.get("timeout"):
        res["error"] += "もう一度「一覧表と照合する」を押してください（一覧表は変わりません）。"
    st.session_state["reconcile_result"] = res

res = st.session_state.get("reconcile_result")
if res:
    if not res.get("ok"):
        st.error("照合できませんでした：%s" % common.md_escape(res.get("error", "不明なエラー")))
    else:
        s = res.get("summary", {})
        diff = s.get("請求書に無い", 0) + s.get("数量違い", 0) + s.get("一覧表に無い", 0) + s.get("noInvoice", 0)
        with st.container(key="metrics"):  # key はスマホ用CSSの目印（.st-key-metrics）
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("請求書に無い", s.get("請求書に無い", 0), help="一覧表にある出荷が請求書に無い＝請求漏れの候補")
            m2.metric("数量違い", s.get("数量違い", 0))
            m3.metric("一覧表に無い", s.get("一覧表に無い", 0))
            m4.metric("請求書が無い得意先", s.get("noInvoice", 0))
        if diff:
            st.warning("ズレがあります。結果のファイルで、色の付いた行を確かめてください（請求書 %d 件・一致 %d・照合対象外 %d）。"
                       % (s.get("invoices", 0), s.get("一致", 0), s.get("照合対象外", 0)))
        else:
            st.success("ズレはありませんでした（請求書 %d 件・一致 %d・照合対象外 %d）。"
                       % (s.get("invoices", 0), s.get("一致", 0), s.get("照合対象外", 0)))
        if s.get("unassigned"):
            st.warning("一覧表の「未振り分け」に、この期間の行が %d 行あります（どの得意先の照合にも入っていません）。"
                       "一覧表で正しい得意先へ移してから、もう一度照合してください。" % s["unassigned"])
        c1, c2 = st.columns([4, 1])
        c1.caption("結果のファイル: %s（GAS %s）" % (common.md_escape(res.get("name", "")), common.md_escape(res.get("codeVersion", ""))))
        if str(res.get("url", "")).startswith("https://docs.google.com/"):
            c2.link_button("結果を開く", res["url"])
        if res.get("customers"):
            st.markdown("#### 得意先ごとの判定")
            st.table(pd.DataFrame(res["customers"]).style.hide(axis="index"))
        if res.get("noInvoice"):
            st.markdown("#### 請求書が無い得意先（同じ締めで、期間内に出荷がある）")
            st.table(pd.DataFrame(res["noInvoice"]).style.hide(axis="index"))
