# DM施策パフォーマンス ダッシュボード

ショップジャパン DM施策の月次KPIを可視化するWebアプリです。
GitHub Pages で無料ホスティングし、チーム全員が同一URLで閲覧できます。

---

## 🗂 ファイル構成

```
dm-dashboard/
├── index.html        ← ダッシュボード本体（これを GitHub Pages で公開）
├── data/
│   └── dm_data.json  ← 施策データ（毎月 convert.py で更新）
├── convert.py        ← xlsm → JSON 変換スクリプト
└── README.md         ← このファイル
```

---

## 🚀 初回セットアップ（15分）

### 1. GitHubリポジトリを作成

1. [github.com](https://github.com) にサインイン
2. 右上の「+」→「New repository」
3. リポジトリ名: `dm-dashboard`（任意）
4. **Public**（GitHub Pages無料プランはPublicのみ）
5. 「Create repository」をクリック

### 2. ファイルをアップロード

```bash
# ローカルのターミナルで実行（Gitがインストール済みの場合）
git init
git add .
git commit -m "initial commit"
git branch -M main
git remote add origin https://github.com/あなたのID/dm-dashboard.git
git push -u origin main
```

または GitHub の「Upload files」ボタンから手動アップロードもOKです。

### 3. GitHub Pages を有効化

1. リポジトリページ →「Settings」タブ
2. 左メニュー「Pages」
3. Source: **Deploy from a branch**
4. Branch: **main** / **/ (root)**
5. 「Save」をクリック

数分後、以下のURLでアクセスできます：
```
https://あなたのGitHubID.github.io/dm-dashboard/
```

---

## 🔄 毎月の更新手順（5分）

### ステップ1: xlsmを最新化（今まで通り）

Excelファイルを更新します（新しい月のシートを追加）。

### ステップ2: 変換スクリプト実行

```bash
# Python 3.8以上が必要
pip install openpyxl   # 初回のみ

python convert.py "DM施策別進捗一覧6.xlsm"
```

実行すると `data/dm_data.json` が自動更新されます。

**新しい月シートを追加した場合**は `convert.py` の `SHEET_LAYOUTS` に追記してください：

```python
SHEET_LAYOUTS = {
    ...
    '5': ('5月詳細', 13, 14, 15, 16, 17, 18, 19, 20, 21),  # ← 追加
}
```

### ステップ3: GitHubにプッシュ

```bash
git add data/dm_data.json
git commit -m "update: 5月データ追加"
git push
```

または GitHub の「Upload files」から `data/dm_data.json` だけアップロードしてもOKです。

---

## 📊 カテゴリ正規化ルール

施策名からカテゴリを自動判別します。`convert.py` の `normalize_category()` 関数で管理しています。

| カテゴリ名 | 判別キーワード |
|---|---|
| お誕生日DM | お誕生日、誕生日 |
| TRS下取定期 | TRS下取（WOW・新規を除く） |
| TRS下取WOW | TRS下取WOW |
| TRS下取新規 | TRS下取新規 |
| TRSキャンペーン | TRSキャンペーン、TRSキャン |
| RAH買い替え | RAH買い替え |
| RAHシリーズ | RAH（買い替え以外） |
| CCH | CCH（シリーズ・買い替え以外） |
| CCH買い替え | CCH買い替え、CCH買替 |
| CCHシリーズ | CCHシリーズ |
| TR7買い替え | TR7買い替え、TR7買替 |
| シニア訴求 | シニア |
| その他 | 上記以外 |

新しい施策カテゴリを追加する場合は `normalize_category()` に条件を追記してください。

---

## 🛠 トラブルシューティング

**Q: `convert.py` でエラーが出る**  
A: 月シートの列構造が変わった可能性があります。`SHEET_LAYOUTS` の列番号を確認してください（0始まり）。

**Q: GitHub Pages が表示されない**  
A: Settings → Pages でデプロイが完了しているか確認してください（初回は数分かかります）。

**Q: データが更新されない**  
A: ブラウザキャッシュをクリア（Ctrl+Shift+R）してください。

---

## 💻 ローカルで動かす場合

`index.html` をダブルクリックしても CORS制限でデータを読めません。
以下のいずれかで起動してください：

```bash
# Python がある場合
python -m http.server 8080
# → http://localhost:8080 でアクセス
```
