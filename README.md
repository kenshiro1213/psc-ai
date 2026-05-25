# psc-ai

ピーエス・コンストラクション株式会社（PSC）様向け AIコンサルティング資料。

## 目的

PSC様に対する明日のヒアリング（60分）の構成・コンテンツ・デモシナリオを整理し、
当日使うHTMLスライド（投影＋PDF配布兼用）と運用ドキュメントを束ねる。

## ドキュメント

| # | ファイル | 内容 |
|---|---|---|
| 00 | [docs/00-概要.md](./docs/00-概要.md) | クライアント情報・目的・制約・事前質問の分類 |
| 01 | [docs/01-ヒアリング構成.md](./docs/01-ヒアリング構成.md) | 60分のタイムテーブル／教育・ヒアリング・クロージング台本 |
| 02 | [docs/02-セキュリティ解説.md](./docs/02-セキュリティ解説.md) | 学習リスク・オプトアウト・インシデント事例・運用パターン |
| 03 | [docs/03-デモシナリオ.md](./docs/03-デモシナリオ.md) | Demo A/B/C の詳細台本・ダミーデータ仕様・プロンプト例 |
| 04 | [docs/04-スライド構成.md](./docs/04-スライド構成.md) | reveal.js想定の全21枚構成／PSCカラーパレット |
| 05 | [docs/05-事前準備.md](./docs/05-事前準備.md) | 当日までのチェックリスト・想定FAQ・リスク対策 |
| 06 | [docs/06-業界事例.md](./docs/06-業界事例.md) | ゼネコン業界のAI活用事例（設計・施工・安全・維持管理・企業内知）。教育パート用 |
| -- | [docs/prompts/](./docs/prompts/) | デモで Claude for Work に投入するプロンプト一式 |
| -- | [demo/](./demo/) | デモ用サンプルデータ（Excel/PDF）と再生成スクリプト |
| -- | [slides/](./slides/) | reveal.js で実装した本番スライド（21枚） |

## クライアントから受領した事前質問

[事前質問.md](./事前質問.md) ── 建設業界の業務効率化に関する12件。

---

## スライドの起動方法

本番スライドは `slides/` 配下に reveal.js で実装済み（全21枚）。
ローカルで起動するにはWebサーバが必要です（`file://` で開くとCSSやロゴが読み込めません）。

### 1. ローカルサーバを起動

リポジトリのルートで：

```bash
cd slides
python3 -m http.server 8765
```

> ポートはお好みで（8000、8080 などでもOK）。

### 2. ブラウザで開く

```bash
open http://localhost:8765/
```

または Chrome / Safari でURLを直接入力。

### 3. 操作キー

| キー | 動作 |
|---|---|
| `→` / `Space` | 次のスライドへ |
| `←` | 前のスライドへ |
| `Esc` | スライド一覧（オーバービュー） |
| `F` | フルスクリーン |
| `S` | スピーカーノート（別ウィンドウ） |
| `?` | キー一覧表示 |

### 4. PDF出力（配布用）

サーバ起動中に、Chromeヘッドレスで一発生成：

```bash
cd slides
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
  --headless --disable-gpu --no-pdf-header-footer \
  --virtual-time-budget=10000 \
  --print-to-pdf="psc-ai-slides.pdf" \
  "http://localhost:8765/?print-pdf"
```

> `?print-pdf` クエリで reveal.js がPDF用レイアウトに切り替わります。
> 出力は `slides/psc-ai-slides.pdf`（21ページ、約2.7MB）。

ブラウザから手動で出力する場合は、`http://localhost:8765/?print-pdf` を開いてから Chrome のメニュー「印刷」→「PDFに保存」（用紙：A4横／余白：なし／背景画像：あり）。

### 5. サーバを止める

`Ctrl+C` で停止。バックグラウンド起動した場合は：

```bash
lsof -ti :8765 | xargs kill
```

---

## 当日までのチェックリスト

詳細は [docs/05-事前準備.md](./docs/05-事前準備.md) を参照。

- [ ] スライドの表示確認（ローカルサーバ）
- [ ] PDF出力確認（配布用）
- [ ] デモ3本のリハーサル（[demo/README.md](./demo/README.md)）
- [ ] 各デモの成功時録画を保存
- [ ] プロジェクタ・ネットワーク・電源の事前確認
