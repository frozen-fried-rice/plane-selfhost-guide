# コピペ用（Attribute 行なし）

`src/*.bas` の先頭には `Attribute VB_Name = "..."` 行があります。これは
**「ファイル→ファイルのインポート」専用**の行で、コード画面に直接貼り付けると
`構文エラー` になります。

このフォルダの `.txt` は、その `Attribute` 行を**取り除いた**コピペ用です。

## 使い方（コピペ派）
1. VBE（`Alt`+`F11`）で **挿入 → 標準モジュール** を2つ作る
2. それぞれに `mQRCode.txt` と `mQRApp.txt` の中身を丸ごと貼り付ける
3. Excel に戻り `Alt`+`F8` → **`QR_Setup`** を実行

> インポート派の人は `src/mQRCode.bas` / `src/mQRApp.bas` を
> 「ファイル→ファイルのインポート」で読み込んでください（`Attribute` 行はそのままでOK）。
