# 出荷連絡表 自動生成ツール

仮納品書（Excel）と商品マスタをアップロードすると、**単価入りの出荷連絡表**を自動生成するWebアプリ（Streamlit）。

## できること
- 仮納品書の全タブを読み、**製品×ロットで数量を合算**（分納を集約）
- **資材／残資材を自動除外**、有償サンプルは含む
- 受領書ブロックは**二重計上しない**
- 商品マスタから**単価を自動ひも付け**（品名の正規化突合＋数量帯＝ロット下限閾値）
- 単価が引けない・数量帯外・名前衝突は **「要確認」** で色付け表示

## ローカル起動
```bash
python3 -m venv .venv
./.venv/bin/pip install -r requirements.txt
./.venv/bin/streamlit run app.py
```
→ http://localhost:8501

## パスワード保護
- **Streamlit Community Cloud**：アプリの Settings → Secrets に以下を設定
  ```toml
  APP_PASSWORD = "好きなパスワード"
  ```
- **ローカル**：`.streamlit/secrets.toml` に同様に記載（未設定なら認証なしで起動）

## デプロイ（Streamlit Community Cloud）
1. https://share.streamlit.io にGitHubでログイン
2. **New app** → このリポジトリ／ブランチ `main`／ファイル `app.py` を選択
3. **Advanced settings → Secrets** に `APP_PASSWORD` を設定
4. **Deploy**

## ファイル
| ファイル | 役割 |
|---|---|
| `engine.py` | 変換ロジック（解析・集計・単価突合・xlsx出力） |
| `app.py` | Streamlit Web UI（パスワード保護付き） |

## ルール概要
- 数量帯：ロット欄の下限数値を閾値に「数量以下で最大の閾値」を採用。最小未満は最小単価＋要確認。
- 価格履歴：同条件で複数あれば最新の価格更新日を採用。
- 名前衝突（正規化後に別商品コード複数）：要確認。
