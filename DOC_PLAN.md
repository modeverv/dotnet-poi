# Phase 9 — Documentation Site Generation Plan

> このドキュメントは「このライブラリを初めて見た開発者」の視点で、
> 「何を知りたいか」「どんな順番で知りたいか」に沿って構成した Phase 9 の実行計画です。

---

## 利用者のペルソナと情報ニーズ

### Persona A: 「とりあえず xlsx を生成したい」
- NuGet の入れ方
- 最小のコードで xlsx を書く方法
- セルに文字列・数値を入れて保存するまで

### Persona B: 「既存の Excel ファイルを読んで加工したい」
- 読み込み API
- セルの値・型・スタイルの取得
- 書き戻し (round-trip)

### Persona C: 「Word / PowerPoint も扱いたい」
- docx の段落・表・画像
- pptx のスライド・テキストボックス・画像・表
- 各フォーマットの対応範囲一覧

### Persona D: 「プロダクション投入を検討している」
- 互換性マトリクス（何ができて何ができないか）
- 制限事項・既知の問題
- パッケージ分割の意味（Core vs Formula）
- Java POI との相互運用性

---

## コンテンツ構成（利用者の情報探索順）

```
docs_src/content/
├── getting-started/          # Persona A が真っ先に見る
│   ├── installation.md       # NuGet のインストール
│   ├── first-workbook.md     # 最初の xlsx 作成
│   ├── first-document.md     # 最初の docx 作成
│   └── first-presentation.md # 最初の pptx 作成
│
├── guides/                   # Persona B, C が次に見る
│   ├── xlsx/
│   │   ├── cell-types.md          # 文字列・数値・日付・論理値・数式
│   │   ├── styles.md              # フォント・塗り潰し・罫線・表示形式・配置
│   │   ├── layout.md              # セル結合・列幅・行高・非表示・固定ペイン
│   │   ├── images.md              # 画像埋め込み・アンカー・回転
│   │   ├── formulas.md            # 数式テキスト書き込み・キャッシュ値
│   │   ├── data-validation.md     # 入力規則
│   │   ├── conditional-formatting.md # 条件付き書式
│   │   ├── auto-filter.md         # オートフィルター
│   │   ├── pivot-tables.md        # ピボットテーブル（プログラム作成）
│   │   ├── protection.md          # ブック保護・セル保護
│   │   ├── rich-text.md           # リッチテキスト（文字ごとの書式）
│   │   └── macros.md              # xlsm マクロ保存
│   │
│   ├── xls/
│   │   └── overview.md            # HSSF の現状と制限
│   │
│   ├── docx/
│   │   ├── paragraphs.md          # 段落・ラン・書式・箇条書き
│   │   ├── tables.md              # 表の作成・読み取り
│   │   ├── images.md              # 画像埋め込み
│   │   ├── headers-footers.md     # ヘッダー・フッター
│   │   ├── hyperlinks.md          # ハイパーリンク
│   │   ├── fields.md              # TOC・ページ番号・差し込み印刷
│   │   └── sections.md            # ページ設定・段組
│   │
│   └── pptx/
│       ├── slides.md              # スライド作成・テキストボックス
│       ├── images.md              # 画像埋め込み・回転
│       ├── tables.md              # 表の作成
│       └── formatting.md          # ラン書式・配色
│
├── compatibility/             # Persona D が確認する
│   ├── format-coverage.md     # フォーマット別対応状況（NOW.md 由来）
│   ├── limitations.md         # 既知の制限・未実装リスト
│   ├── interop.md             # Java POI との相互運用
│   └── package-split.md       # Core vs Formula の分割理由
│
└── reference/                 # 必要に応じて（Phase 9 では軽めに）
    └── examples-index.md      # examples/ 以下の全プロジェクト一覧
```

---

## 各ページと対応する example コード

| ドキュメント | 対応 example | 備考 |
|---|---|---|
| `installation.md` | なし（手順説明のみ） | NuGet パッケージ構成、csproj への追加方法 |
| `first-workbook.md` | `Phase0WriteExample` | 最小 xlsx 作成 |
| `first-document.md` | `Phase32DocxExample` | 最小 docx 作成 |
| `first-presentation.md` | `Phase33PptxExample` | 最小 pptx 作成 |
| `xlsx/cell-types.md` | `Phase7CellTypesExample` | 全セル型の読み書き |
| `xlsx/styles.md` | `UsageSamples` (CreateSpreadsheet) | フォント・塗り潰し・罫線・表示形式 |
| `xlsx/layout.md` | `Phase3InterfaceExample` | セル結合・列幅・行高 |
| `xlsx/images.md` | `Phase25ImagesExample` | 画像埋め込み |
| `xlsx/formulas.md` | `Phase5FormulaEvaluatorExample` | 数式テキストと評価 |
| `xlsx/data-validation.md` | `UsageSamples` (CreateSpreadsheet) | 入力規則 |
| `xlsx/auto-filter.md` | 新規作成候補 | オートフィルター |
| `xlsx/pivot-tables.md` | 追加検討 | ピボットテーブル |
| `xlsx/protection.md` | 追加検討 | シート保護・ブック保護 |
| `xlsx/rich-text.md` | 追加検討 | リッチテキスト |
| `xlsx/macros.md` | 追加検討 | xlsm マクロ保存 |
| `docx/paragraphs.md` | `Phase32DocxExample` | 段落・書式 |
| `docx/tables.md` | `UsageSamples` (CreateDocument) | 表 |
| `docx/images.md` | `UsageSamples` (CreateDocument) | 画像 |
| `docx/headers-footers.md` | 新規作成候補 | ヘッダー・フッター |
| `docx/hyperlinks.md` | 新規作成候補 | ハイパーリンク |
| `docx/fields.md` | 新規作成候補 | TOC・ページ番号 |
| `docx/sections.md` | 新規作成候補 | ページ設定 |
| `pptx/slides.md` | `Phase33PptxExample` | スライド・テキストボックス |
| `pptx/images.md` | `Phase33PptxExample` | 画像 |
| `pptx/tables.md` | `UsageSamples` (CreatePresentation) | 表 |
| `pptx/formatting.md` | `Phase33PptxExample` | ラン書式 |
| `compatibility/format-coverage.md` | なし（NOW.md 由来） | フォーマット別対応状況 |
| `compatibility/limitations.md` | なし（NOW.md 由来） | 未実装・制限の一覧 |
| `compatibility/interop.md` | `Phase1InteropExample` | Java POI 相互運用 |
| `compatibility/package-split.md` | なし（設計説明） | Core vs Formula |
| `reference/examples-index.md` | 全 example | 各 example の概要と実行方法 |

---

## 実装ステップ

### Step 1: 既存 example の検証と不足 example の作成

各 example プロジェクトを `dotnet run` で実行し、正しく動作することを確認する。その上で、以下のドキュメントに対応する example がないものについて、不足を判断する。

**要判断（Phase 9 内で作るかどうか）：**
- AutoFilter example（現時点で example なし → ガイドにスニペットを埋め込むか、example を作るか）
- PivotTable example（同上）
- Protection example（同上）
- RichText example（同上）
- Macro example（同上）
- Headers/Footers example（同上）
- Hyperlinks example（同上）
- Fields example（同上）
- Sections example（同上）

**判断基準：**
- ガイドに載せるスニペット（20行以内）で説明が完結するもの → example は作らずスニペットを直接 markdown に書く
- 複数ファイルや長いコードが必要なもの → `examples/` にプロジェクトを作る

### Step 2: docs_src/ の Markdown 執筆

**ルール：**
- 各ページは **1 利用者タスク = 1 ページ** を原則とする
- ページ冒頭に **コードスニペット（20行以内）** を置き、すぐに動かせるイメージを持たせる
- 詳細説明はスニペットの後
- スニペット直後またはページ末尾に `examples/` の該当プロジェクトへのリンクを置く
- Apache POI 文書からの長文コピーはしない。dotnet-poi の言葉で書く
- パッケージ分割（Core vs Formula）に言及する場合は正確に。数式評価の対応範囲を拡大解釈させない

**執筆優先順位（利用者が知りたい順）：**
1. **getting-started/installation.md** — NuGet の入れ方が分からなければ何も始まらない
2. **getting-started/first-workbook.md** — 最初に触るのは xlsx
3. **getting-started/first-document.md** — docx
4. **getting-started/first-presentation.md** — pptx
5. **compatibility/format-coverage.md** — 何がどこまでできるか
6. **compatibility/limitations.md** — 何ができないか（明確に）
7. **compatibility/package-split.md** — なぜ 2 パッケージなのか
8. **compatibility/interop.md** — Java POI との関係
9. 各ガイド（xlsx → docx → pptx の順）

### Step 3: DocsGenerator の機能改善

現在の DocsGenerator は最小限の Markdown→HTML 変換を持っている。必要に応じて以下を追加する：

- ~~テーブル記法対応~~（当面必要なら Markdig 導入を検討）
- コードブロックのコピーボタン（CSS のみで実現可能なら）
- 目次（TOC）の自動生成（ページ内見出しからのリンク）
- シンタックスハイライト（CSS クラスを振る + 利用者ブラウザで prism.js 等を CDN 読み込み）

**ただし、必要になるまで機能追加しない。** 最初は今のままで十分。

### Step 4: HTML 生成と spot-check

```bash
dotnet run --project tools/DotnetPoi.DocsGenerator -- docs_src docs
```

生成後、以下の観点で spot-check する：
- ナビゲーションが全てのページをカバーしているか
- リンクが切れていないか
- コードブロックが正しく表示されているか
- モバイルレイアウトが崩れていないか

### Step 5: CHECKPOINT.md の更新

Phase 9 完了時に以下を記録する：
- 作成した docs_src ページ数
- 作成・更新した example 数
- 生成された docs/ のファイル数
- 残課題（書けなかったページ・将来追加したいコンテンツ）

---

## スケジュール目安

| Step | 内容 | 目安 |
|---|---|---|
| 0 | 現状の example 全件実行・動作確認 | — |
| 1 | 不足 example の作成判断・実装 | 要相談 |
| 2 | getting-started (4 pages) | — |
| 3 | compatibility (4 pages) | — |
| 4 | xlsx guides (12 pages) | — |
| 5 | docx guides (7 pages) | — |
| 6 | pptx guides (4 pages) | — |
| 7 | reference/examples-index.md | — |
| 8 | DocsGenerator 機能改善（必要な場合） | — |
| 9 | HTML 生成・spot-check・微調整 | — |
| 10 | CHECKPOINT 更新 | — |

---

## 「やらないこと」リスト（Phase 9 のスコープ外）

- API リファレンスの自動生成（XmlDoc → HTML）
- クラス階層図・アーキテクチャ解説
- Apache POI からの機械翻訳
- 多言語対応（英語のみ）
- 全文検索機能
- docs/ の手書き編集（必ず docs_src → 生成）
- テストの新規追加（既存 example の検証は行うが、テスト不足の補填は Phase 9 の目的ではない）
