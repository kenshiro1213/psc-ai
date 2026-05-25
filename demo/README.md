# デモ用サンプルデータ

明日のヒアリングデモで使うサンプルデータ一式。すべて **架空のダミーデータ** です。

## ツール

すべて **Claude for Work（Claude.ai のチャットUI）** を使用。
- 各デモの Excel ファイルをチャットに添付
- [docs/prompts/](../docs/prompts/) のプロンプトを順に投入
- 出力を確認・追加プロンプトで深掘り

## ディレクトリ構成

```
demo/
├── README.md                              # 本ファイル
├── scripts/                               # Excel 再生成スクリプト（openpyxl）
│   ├── build_demo_a.py
│   ├── build_demo_b.py
│   └── build_demo_c.py
├── demo-a-handover/
│   └── 顧客A_引継ぎ資料.xlsx              # 5シート（案件概要/メール/議事録/要望/キーパーソン）
├── demo-b-docgen/
│   ├── 見積要綱書_サンプル.pdf            # 入力（Chrome印刷で.htmlから生成）
│   ├── 見積要綱書_サンプル.html           # PDF再生成用ソース
│   ├── 見積要綱書_サンプル.md             # 内容の可読版
│   └── 社内図説シート_テンプレート.xlsx   # 3シート（図説/記入要領/凡例）
└── demo-c-bid-watch/
    └── 発注公告_収集データ.xlsx           # 3シート（公告データ30件/抽出条件/出力テンプレ）
```

## 再生成方法

データを編集したいときは `scripts/build_demo_*.py` を編集してから：

```bash
cd psc-ai
python3 demo/scripts/build_demo_a.py
python3 demo/scripts/build_demo_b.py
python3 demo/scripts/build_demo_c.py
```

PDFの再生成（Chrome必須）：
```bash
cd demo/demo-b-docgen
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
  --headless --disable-gpu --no-pdf-header-footer \
  --print-to-pdf="見積要綱書_サンプル.pdf" \
  "file://$(pwd)/見積要綱書_サンプル.html"
```

## デモ毎の運用

### Demo A（顧客引継ぎ）
1. Claude for Work に `demo-a-handover/顧客A_引継ぎ資料.xlsx` を添付
2. `docs/prompts/demo-a-prompt.md` のプロンプト1を投入
3. プロンプト2（質問リスト）を追加投入

### Demo B（仕様書→ドキュメント）
1. Claude for Work に `見積要綱書_サンプル.pdf` と `社内図説シート_テンプレート.xlsx` を添付
2. `docs/prompts/demo-b-prompt.md` のプロンプト1〜3を順に投入

### Demo C（発注公告 絞り込み）
1. Claude for Work に `発注公告_収集データ.xlsx` を添付
2. `docs/prompts/demo-c-prompt.md` のプロンプト1〜3を順に投入

## 機密性メモ

- 全データは **架空** です。実在の自治体・企業を想起させる地名はあくまでサンプルです
- 明日のデモでも「すべてダミーデータです」と冒頭で明言してください
