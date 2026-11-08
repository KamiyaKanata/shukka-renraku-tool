# -*- coding: utf-8 -*-
"""出荷連絡表 自動生成 — Web版。画面は3つ（左のメニューで切り替え）。
- 出荷連絡表を作る（views/shukka.py）: 仮納品書 → 単価入りの出荷連絡表＋得意先ごと出荷一覧表への登録
- 請求書と照合（views/seikyu_shogo.py）: 販売大将の請求書CSV → 一覧表と数量を比べてズレを出す
- 販売大将に取り込む（views/hanbai_import.py）: 一覧表の期間の行 → 表で直す → 売上伝票インポートのCSV
起動: ./.venv/bin/streamlit run app.py
"""
import importlib

import streamlit as st

import common

importlib.reload(common)  # Streamlit Cloudの旧モジュール残留対策

st.set_page_config(page_title="出荷連絡表 自動生成", page_icon="📦", layout="wide")

if not common.require_password():
    st.stop()

st.navigation([
    st.Page("views/shukka.py", title="出荷連絡表を作る", icon="📦", default=True),
    st.Page("views/seikyu_shogo.py", title="請求書と照合", icon="🧾"),
    st.Page("views/hanbai_import.py", title="販売大将に取り込む", icon="📤"),
]).run()
