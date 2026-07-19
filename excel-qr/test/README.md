# 検証について（参考）

`verify_reference.py` は、`src/mQRCode.bas` と **同一のアルゴリズム** を Python に移植したものです。
VBA を実行できない環境でも、生成ロジックの正しさを機械的に確認するために使いました。

確認済みの内容：

- 標準ライブラリ **segno** と、型番選択・データ領域のビット列が一致
- Reed-Solomon 符号語が生成多項式で割り切れる（＝誤り訂正が数学的に正しい）
- 実機系デコーダ **zbar** で、型番 1〜40・誤り訂正 L/M/Q/H・英数字/日本語/URL/長文まで
  復元バイトが元テキストの UTF-8 と完全一致

実行（任意・開発者向け）:

```bash
pip install segno pyzbar pillow numpy
# Linux は libzbar0 も必要:  sudo apt-get install -y libzbar0
python3 verify_reference.py
```

> このスクリプトはあくまで検証用の参考です。QRアプリ本体の動作に Python は不要で、
> Excel の VBA だけで完結します。
