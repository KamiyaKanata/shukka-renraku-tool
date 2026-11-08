# -*- coding: utf-8 -*-
"""2つの画面（出荷連絡表を作る／請求書と照合）で共通の部品。"""
import json

import streamlit as st

APP_VERSION = "v5.6（送り状の読み取り）"

_MD_SPECIAL = "\\`*_{}[]()#+-.!|<>~"


def md_escape(s):
    """GASから返った文字列（シート名・理由など）をマークダウンとして解釈させない。"""
    return "".join("\\" + ch if ch in _MD_SPECIAL else ch for ch in str(s))


def apply_css():
    """スマホ対応（固定のCSSのみ。外から来た文字列は入れない）
    - 表は横にスクロールさせる（画面幅に押し込むと製品名が1文字ずつ改行され、右の列が切れて見えなくなる）
    - スマホ幅では件数の4つを2列×2段にし、タイトルを小さくする"""
    st.markdown("""<style>
[data-testid="stTable"]{overflow-x:auto !important;-webkit-overflow-scrolling:touch}
[data-testid="stTable"] table{min-width:720px}
@media (max-width:640px){
  h1{font-size:1.5rem !important;line-height:1.4 !important}
  .st-key-metrics [data-testid="stHorizontalBlock"]{flex-wrap:wrap !important;gap:.25rem 1rem !important}
  .st-key-metrics [data-testid="stColumn"]{flex:1 1 calc(50% - 1rem) !important;min-width:calc(50% - 1rem) !important;width:calc(50% - 1rem) !important}
  .st-key-metrics [data-testid="stMetricValue"]{font-size:1.6rem !important}
}
</style>""", unsafe_allow_html=True)


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


def post_ichiran(payload, timeout=90):
    """得意先ごと出荷一覧表のGASへ送る（合言葉と版はここで足す）。返り値はGASの応答（dict）。
    失敗も dict（ok=False, error=理由）で返し、例外は外に出さない。"""
    import requests
    try:
        url = st.secrets["ICHIRAN_GAS_URL"]
        token = st.secrets["ICHIRAN_GAS_TOKEN"]
    except Exception:
        return {"ok": False, "error": "一覧表の接続先（Secrets の ICHIRAN_GAS_URL / ICHIRAN_GAS_TOKEN）が未設定です。"}
    body = json.dumps(dict(payload, token=token, appVersion=APP_VERSION), ensure_ascii=False)
    try:
        r = requests.post(url, data=body.encode("utf-8"),
                          headers={"Content-Type": "application/json; charset=utf-8"}, timeout=timeout)
    except requests.Timeout:
        return {"ok": False, "timeout": True, "error": "一覧表からの応答が時間内に返りませんでした。"}
    except requests.RequestException as e:
        return {"ok": False, "error": "一覧表に接続できませんでした（%s）。" % type(e).__name__}
    try:
        data = r.json()
    except ValueError:
        return {"ok": False, "error": "一覧表から想定外の応答がありました（HTTP %s）。GASの公開設定（アクセス: 全員）を確認してください。"
                                      % r.status_code}
    if not isinstance(data, dict):
        return {"ok": False, "error": "一覧表から想定外の応答がありました。"}
    return data
