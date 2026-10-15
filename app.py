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

# スマホ対応（固定のCSSのみ。外から来た文字列は入れない）
# - 表は横にスクロールさせる（画面幅に押し込むと製品名が1文字ずつ改行され、右の列が切れて見えなくなる）
# - スマホ幅では件数の4つを2列×2段にし、タイトルを小さくする
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


@st.cache_data(ttl=1800, show_spinner=False)
def _fetch_master():
    """商品マスタ（ロット別単価マスタ）をGAS Web App経由で取得し、単価索引を返す。30分キャッシュ。"""
    import requests
    url = st.secrets["MASTER_GAS_URL"]
    token = st.secrets["MASTER_GAS_TOKEN"]
    r = requests.get(url, params={"token": token}, timeout=90)
    r.raise_for_status()
    data = r.json()
    if isinstance(data, dict) and data.get("error"):
        raise RuntimeError("マスタAPIエラー: %s（トークン不一致など）" % data["error"])
    return engine.load_master_from_rows(data["values"])


def _detail_preview_df(groups):
    """ダウンロードの『詳細シート』と同じ内容（得意先＝シート名／商品CD昇順／CD0補正）。"""
    dated = [(g["日付"], r) for g in groups for r in g["records"]]
    dated = sorted(dated, key=lambda dr: engine._cd_sortkey(dr[1]["商品CD"]))
    rows = [{
        "得意先": r.get("得意先", ""),
        "製品名": engine.detail_name(r), "ロット": r["ロット"],
        "出荷数": engine._qty_disp(r),
        "単価": ("" if r["単価"] is None else r["単価"]),
        "ケース": engine._cases_str(r["cases"]),
        "商品CD": engine._pad_cd(r["商品CD"]), "試作番号": r["試作番号"],
        "要確認": ("要確認" if r["要確認"] else ""), "メモ": r["メモ"],
    } for dl, r in dated]
    return pd.DataFrame(rows) if rows else pd.DataFrame({"(なし)": []})


_MD_SPECIAL = "\\`*_{}[]()#+-.!|<>~"


def _md_escape(s):
    """GASから返った文字列（シート名・理由など）をマークダウンとして解釈させない。"""
    return "".join("\\" + ch if ch in _MD_SPECIAL else ch for ch in str(s))


def _register_ichiran(rows, files):
    """得意先ごと出荷一覧表のGASへ送る。返り値はGASの応答（dict）。失敗も dict で返し、例外は外に出さない。"""
    import json
    import requests
    try:
        url = st.secrets["ICHIRAN_GAS_URL"]
        token = st.secrets["ICHIRAN_GAS_TOKEN"]
    except Exception:
        return {"ok": False, "error": "一覧表の接続先（Secrets の ICHIRAN_GAS_URL / ICHIRAN_GAS_TOKEN）が未設定のため、登録していません。"}
    body = json.dumps({"token": token, "rows": rows, "files": files, "appVersion": APP_VERSION}, ensure_ascii=False)
    try:
        r = requests.post(url, data=body.encode("utf-8"),
                          headers={"Content-Type": "application/json; charset=utf-8"}, timeout=90)
    except requests.Timeout:
        return {"ok": False, "error": "一覧表からの応答が時間内に返りませんでした。登録できている可能性があります。"
                                      "もう一度「生成」を押しても二重には入りません。"}
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


def _show_ichiran_result(res, latest=False):
    """一覧表への登録結果を表示。登録できなかったときは必ず赤で出す（黙って欠けないように）。"""
    if not res:
        return
    when = "（直前の生成分）" if latest else ""
    if res.get("skipped"):
        st.warning("📒 得意先ごと出荷一覧表%sには登録していません（「得意先ごと出荷一覧表にも登録する」のチェックを外して生成したため）。" % when)
        return
    if not res.get("ok"):
        st.error("📒 得意先ごと出荷一覧表%sに登録できませんでした：%s" % (when, _md_escape(res.get("error", "不明なエラー"))))
        return
    s = res.get("summary", {})
    det = res.get("details", {})
    st.success("📒 得意先ごと出荷一覧表%sに登録しました：新規 %d 行／登録済みのため飛ばした %d 行"
               % (when, s.get("added", 0), s.get("skipped", 0)))
    c1, c2 = st.columns([4, 1])
    c1.caption("登録先: %s（GAS %s）" % (_md_escape(res.get("sheetName", "")), _md_escape(res.get("codeVersion", ""))))
    if str(res.get("url", "")).startswith("https://docs.google.com/"):
        c2.link_button("一覧表を開く", res["url"])
    if s.get("unassigned"):
        st.warning("🔶 %d 行を「未振り分け」に置きました。一覧表で正しい得意先のシートへ行ごと移してください。" % s["unassigned"])
        st.table(pd.DataFrame(det.get("unassigned", [])))
    if s.get("flagged"):
        st.warning("🔶 %d 行が「数量違いの再登録」です。一覧表の「要確認・メモ」を確かめてください。" % s["flagged"])
        st.table(pd.DataFrame(det.get("flagged", [])))
    if det.get("skipped"):
        with st.expander("登録済みのため飛ばした行（%d 行）— 追加の出荷が混ざっていないか確認用" % len(det["skipped"])):
            st.table(pd.DataFrame(det["skipped"]))


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

APP_VERSION = "v5.2（得意先ごと出荷一覧表への登録を選べる）"
st.title("📦 出荷連絡表 自動生成（MVP）")
st.caption(f"仮納品書をアップロードして「生成」を押すと、単価入りの出荷連絡表ができます。単価は商品マスタから自動取得。｜{APP_VERSION}")

with st.sidebar:
    st.header("使い方")
    st.markdown(
        "1. **仮納品書**（Excel）をアップロード（**複数選択可**）\n"
        "2. **生成**ボタンを押す（単価は商品マスタから自動取得・アップロード不要）\n"
        "3. 表を確認して **ダウンロード**\n\n"
        "・複数ファイル／複数シートは **出荷日ごとに別シート**で1つのExcelにまとめます\n"
        "・同一日付の中は製品×ロットで合算します\n"
        "・資材／残資材は自動で除外、有償サンプルは含めます\n"
        "・**赤文字（黒との二重記載の控え）は自動で除外**し二重計上しません\n"
        "・単価は品名＋数量（ロット閾値）で自動ひも付け\n"
        "・引けない／曖昧な行は **🔶要確認** として表示\n"
        "・生成すると **得意先ごと出荷一覧表** にも自動で登録（同じ内容は二重に入りません。入れたくないときは生成ボタンの上のチェックを外す）\n"
        "・直近3件は画面下の **履歴** から再ダウンロード可（このセッション内のみ）"
    )
    date_label = st.text_input("出荷日（空欄なら仮納品書から自動取得）", value="",
                               help="例: 2026-06-08。空欄のままなら仮納品書の日付を自動で使います。")

kari = st.file_uploader("仮納品書（.xlsx / .xlsm）※複数選択可", type=["xlsx", "xlsm"],
                        accept_multiple_files=True)
st.caption("単価は商品マスタ（ロット別単価マスタ）から自動取得します。マスタのアップロードは不要です。")

# 一覧表への登録は初期値オン（付け忘れが請求漏れの元になるため）。昔の日付の印刷し直しなど、入れたくないときだけ外す
register_ichiran = st.checkbox(
    "得意先ごと出荷一覧表にも登録する", value=True, key="register_ichiran",
    help="昔の日付の出荷連絡表を印刷し直すときなど、一覧表に入れたくないときだけ外してください。"
         "チェックを入れたまま同じ仮納品書で2回生成しても、同じ内容は二重には入りません。")
go = st.button("🚀 出荷連絡表を生成", type="primary", disabled=not kari)

if go and kari:
    try:
        with st.spinner("商品マスタを取得中…"):
            master_index, master_count = _fetch_master()
    except Exception as e:
        st.error(f"商品マスタの取得に失敗しました（Secrets の MASTER_GAS_URL / MASTER_GAS_TOKEN をご確認ください）: {e}")
        st.stop()
    try:
        with st.spinner("解析中…"):
            groups, stats, debug = engine.generate(kari, master_index, master_count, date_label)
    except Exception as e:
        st.error(f"処理でエラーが発生しました: {e}")
        st.stop()

    dates = stats["日付一覧"]
    with st.container(key="metrics"):  # key はスマホ用CSSの目印（.st-key-metrics）
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
    # on_click="ignore": ダウンロードで画面を作り直さない（下の一覧表への登録を途中で止めないため）
    st.download_button("⬇️ 出荷連絡表をダウンロード", data=bio.getvalue(), file_name=fname,
                       mime=_XLSX_MIME, type="primary", on_click="ignore")
    summary = f"{len(dates)}シート / {stats['出荷連絡表 行数']}行 / 要確認{stats['要確認 行数']}"
    _push_history(fname, bio.getvalue(), summary)

    # 得意先ごと出荷一覧表への登録。出荷連絡表の作成とは独立させ、失敗しても上のダウンロードはそのまま使える
    try:
        ich_rows, ich_problem = engine.ichiran_rows(groups, debug["items"], date_label)
        if not register_ichiran:
            ich_res = {"ok": False, "skipped": True}
        elif ich_problem:
            ich_res = {"ok": False, "error": ich_problem}
        else:
            with st.spinner("得意先ごと出荷一覧表に登録中…"):
                ich_res = _register_ichiran(ich_rows, [getattr(f, "name", "") for f in kari])
    except Exception as e:  # 一覧表側の不具合で出荷連絡表の画面まで止めない
        ich_res = {"ok": False, "error": "登録の準備中に想定外のエラーが起きました（%s）。" % type(e).__name__}
    st.session_state["ichiran_result"] = ich_res
    _show_ichiran_result(ich_res)

    # プレビュー＝ダウンロードの「詳細シート」と同じ内容。印刷用シートはダウンロードにのみ入る。
    st.markdown("#### 内容プレビュー（＝詳細シートと同じ。ダウンロードには別途『印刷用シート（実物フォーマット）』も入ります）")
    st.caption("出荷日: " + "、".join(f"{d}（{engine._weekday_jp(d)}）" for d in dates)
               + "　※スマホでは表を横にスクロールできます")
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
    st.info("仮納品書をアップロードすると「生成」できます（単価は商品マスタから自動取得）。")
    _show_ichiran_result(st.session_state.get("ichiran_result"), latest=True)

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
