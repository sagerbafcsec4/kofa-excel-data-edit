# Excelデータ加工

サッカー記録ワークブックのレイアウト整形をブラウザ上で行うWebアプリ。
公開URL: https://sagerbafcsec4.github.io/kofa-excel-data-edit/

---

## ⚠ 修正する人へ（最重要）

**ロジックの唯一の正本は `format_records.py` です。**
`index.html` は実行時に `format_records.py` を読み込んで動くので、
**ロジックの修正は `format_records.py` だけ直せばOK**（index.html 側にコードを書き写す必要はありません）。

修正手順:
1. `format_records.py` を編集
2. このリポジトリにアップロード（コミット）
3. 数十秒で GitHub Pages に自動反映（古い表示が出たら Ctrl+F5）

別のPC・スマホ、別のClaude/Coworkで作業するときも、
**まずこの README と `format_records.py` を読めば全体を把握できます。**

---

## 絶対ルール（データ保護）

- セルの値・数字は**書き換えない**。行う加工は「行/列/シートの削除」「列幅/フォント/縮小表示/印刷範囲の調整」のみ。
- **例外**: 「試合日」を指定したときだけ、表記シートの「歳」列に年齢式 `=DATEDIF(生年月日, "試合日", "Y")` を入れ、
  監督（最終データ行）の年齢セルの直下に注記「※年齢は試合当日の年齢」を入れる。
- 出力前に**全セル値を自動照合**（`process_checked` → `verify_no_value_change`）。
  値の変更/追加を検知したら**出力を中止**。歳列は試合日を指定したときのみ照合対象外。
- 元ファイルは変更せず、別ファイル（`_processed` 付き）をダウンロードする方式。

## シート判定（名前ではなく構造で判定）

- §A 出場記録 / §B シーズン記録 / §C フォーメーション(過去) / §D 表記 / §E 経歴
- シート名を変えても、ヘッダーの中身で自動判別する。

## ファイル一覧

| ファイル | 役割 |
|---|---|
| `index.html` | Webアプリ。実行時に `format_records.py` を読み込む |
| `format_records.py` | **整形ロジック本体（唯一の正本）** |
| `soccer_format_app.pyw` / `build_exe.py` / `EXEを作成.bat` / `アプリを起動.bat` | デスクトップ版（インストール不要の.exe作成用） |
| `はじめにお読みください.md` / `README_公開手順.md` | 説明書 |

## 仕組み

- ブラウザ上の Python（Pyodide）＋ openpyxl で `format_records.py` を実行。
- 処理はすべて利用者のブラウザ内で完結し、ファイルは外部送信されない。
