# Excelデータ加工（Webアプリ版）公開手順

`index.html` 1ファイルだけで動く完全クライアント型アプリです。お手持ちの
kofa-word-format / kofa-excel-format と同じく **GitHub Pages** で公開できます。
（サーバー不要・整形はすべて利用者のブラウザ内で実行。ファイルは外部送信されません）

## 使い方（利用者）
1. 公開URLを開く（初回はエンジン準備に数十秒、以降はキャッシュで高速）。
2. `.xlsx` をドラッグ＆ドロップ、または選択（複数可）。
3. 「経歴シートを削除する」を必要に応じて調整。
4. 「▶ 実行」→ 各ファイルのダウンロード、または「ZIPでまとめてダウンロード」。

## GitHub Pages で公開する手順
1. GitHub で新規リポジトリを作成（例: `kofa-excel-edit`）。
2. この `index.html` をアップロード（ドラッグ＆ドロップでOK）。
3. リポジトリの **Settings → Pages** で
   「Branch: main / (root)」を選んで Save。
4. 数十秒後、`https://<ユーザー名>.github.io/<リポジトリ名>/` で公開。

## 自分のPCで先に試す場合
- `index.html` をブラウザにドラッグして開く方法は、ブラウザによっては
  Pyodide が動かないことがあります。確実なのは GitHub Pages 上での確認です。
- Python があるなら、このフォルダで `python -m http.server` を実行し
  `http://localhost:8000/` を開くと確認できます。

## 仕組み
- ブラウザ上の Python（Pyodide）＋ openpyxl で `format_records.py` のロジックを
  そのまま実行します。シートは名前ではなく**構造**で自動判別（§A〜§E）。
- 元ファイルは読み込むだけ。整形済みコピーをダウンロードする方式なので安全です。
