# CHECKPOINT

## 2026-05-06 xx:xx JST - Refresh README status and structure

- Current task: ルート `README.md` の Status を `NOW.md` ベースの現在カバレッジに差し替え、古い Phase -1/0 などの詳細履歴を README から除去し、プロジェクト構造を最新化する。
- Scope: README と CHECKPOINT の更新。コミットしない。

### やったこと

- `README.md` の `## Status` 以下を、Phase 履歴ではなく `NOW.md` ベースの format coverage / practical gaps / test snapshot に差し替えた。
- README から古い `Phase -1` / `Phase 0` / 個別 phase verification の長い記述を削除した。
- Practical Gaps の comment 項目を、既存コメントは unknown-part preservation で round-trip 保持される想定だが API read/create/edit 未対応、という表現に修正した。
- Runnable examples 一覧に `Phase4HssfXlsExample`、`Phase7CellTypesExample`、`Phase8CoreOnlyExample`、`UsageSamples` を追加した。
- Repository Structure に `docs_src/`、`docs/`、`tools/DotnetPoi.DocsGenerator/`、`tools/XmlCheck/`、`tools/test.sh`、`NOW.md`、`README.jp.md` を反映した。
- Repository Structure を現物の `find` 結果に合わせて再更新し、`.github/`、`src/` 配下の module directories、`tests/` fixtures、`examples/`、`docs_src/`、`docs/`、`tools/`、root files まで明示した。

### Verification

- `rg` で README 内の旧テストプロジェクト参照（`DotnetPoi.XSSF.Tests` など）や `Phase -1` / `Phase 0` の残存がないことを確認。
- `git diff -- README.md CHECKPOINT.md` で差分確認。
- テストは未実行。直前の確認では `dotnet test tests/DotnetPoi.Core.Tests/ --no-restore` が build で失敗:
  - `HSSFSheet` missing `ISheet.setActiveCell/getActiveCell/setSelected/isSelected`
  - `HSSFWorkbook` missing `IWorkbook.setActiveSheet/getActiveSheetIndex/setSelectedTab`

## 2026-05-06 xx:xx JST - Move devbox files under tools/dev

- Current task: `docker-compose.yml`、`.env.sample`、`Dockerfile.devbox` を `tools/dev/` に移動し、devbox が起動後も動き続けるように修正する。
- Scope: devbox 設定ファイルと README 構成図の更新。コミットしない。

### やったこと

- `tools/dev/` を作成。
- `docker-compose.yml` → `tools/dev/docker-compose.yml` に移動。
- `.env.sample` → `tools/dev/.env.sample` に移動。
- `Dockerfile.devbox` → `tools/dev/Dockerfile.devbox` に移動。
- compose の build context / Dockerfile path / workspace volume を、`tools/dev/` 配置からリポジトリルートを指すように変更。
- compose に `init: true` と `exec tail -f /dev/null` を追加し、設定生成後もコンテナが動き続けるようにした。
- Dockerfile に `CMD ["tail", "-f", "/dev/null"]` を追加し、image 単体起動でも落ちないようにした。
- `.env.sample` に `DEEPSEEK_BASE_URL` と `DEEPSEEK_MODEL` を追加。
- README の Repository Structure を `tools/dev/` 配置に更新し、ルートの `docker-compose.yml` / `Dockerfile.devbox` 表記を削除。

### Verification

- `docker compose -f tools/dev/docker-compose.yml config` が成功し、build context と Dockerfile path がリポジトリルート基準で解決されることを確認。
- Docker build/up はネットワーク・イメージ取得が必要になる可能性があるため未実行。

## 2026-05-05 13:xx JST - Add Phase 9 documentation generation guidance

- Current task: POI ドキュメントを参考にしつつ、dotnet-poi 独自の Markdown + runnable examples + HTML 生成を進める方針を `agents.md` に Phase 9 として追加。
- Follow-up: ドキュメントソースと生成スクリプト/設定の標準配置場所を `docs_src/` に明示。
- Scope: コミットしない。

### やったこと

- `agents.md` の英語版フェーズ一覧に `Phase 9 — Documentation Site Generation` を追加。
- `agents.md` の日本語版フェーズ一覧にも同内容を追加。
- POI 文書は Apache License 2.0 の参考資料として読み、長文コピーや構造の丸写しを避ける方針を明記。
- `examples/` の実コードを先に作って検証し、その結果を Markdown docs に反映し、HTML を `docs/` に生成するワークフローを明記。
- Markdown 原稿と生成スクリプト/設定は `docs_src/`、生成 HTML は `docs/`、runnable example code は `examples/` に置くように追記。
- `dotnet` / Java / Python など、再現可能であれば任意のツールを docs 生成・検証に使ってよいことを明記。

### Verification

- `git diff -- agents.md CHECKPOINT.md` で追記内容を確認。
- ドキュメント更新のみなのでテストは未実行。

## 2026-05-05 13:xx JST - Add Phase 7.1 interop verification gate

- Current task: agents.md に interop 検証用のフェーズを追加してチェック項目を明文化。
- Scope: コミットしない。

### やったこと

- `agents.md` に Phase 7.1 を追加（英日両方）。
- interop の必須チェック項目（方向A/B、fixture運用、コア項目、マクロ、未知パーツ保持、XMLパリティ条件）を列挙。

### Verification

- 未実施（ドキュメント更新のみ）。

## 2026-05-05 12:xx JST - Added local interop test script

- Current task: CI と同じ Java↔.NET interop テストをローカル実行する `test.sh` を追加。
- Scope: コミットしない。

### やったこと

- `test.sh` をリポジトリ直下に追加。
  - `dotnet build DotnetPOI.sln` 実行。
  - 方向A: `WriteForDotnetTest` → C# `Category=ReadFromPoi`。
  - 方向B: C# `Category=WriteForPoi` → `ReadFromDotnetTest`。
  - `dotnet` / `mvn` / `java` の存在チェックを追加。

### Verification

- `./test.sh` を実行。
  - .NET ビルド成功（NU1603 警告あり）。
  - Maven テスト: WriteForDotnetTest 4/4、ReadFromDotnetTest 16/16 成功。
  - C# interop テスト: ReadFromPoi 4/4、WriteForPoi 16/16 成功。

## 2026-05-05 12:xx JST - Project restructured: Core + Formula

- Current task: `src/` 配下の全プロジェクトを **DotnetPoi.Core** と **DotnetPoi.Formula** の 2 プロジェクトに統合し、それぞれ独立した NuGet としてビルド可能にした。
- Scope: コミットしない。

### やったこと

**プロジェクト構成変更:**

- 旧 8 プロジェクト (SS, POIFS, XSSF, HSSF, HWPF, XWPF, XSLF, HSLF) の `.csproj` を削除。
- 代わりに `src/DotnetPoi.Core/DotnetPoi.Core.csproj` を作成。
  - `<Compile Include="...">` で旧プロジェクトの全 `.cs` ファイルを単一アセンブリ `DotnetPoi.Core.dll` にコンパイル。
  - 名前空間は従来通り（`DotnetPoi.XSSF.UserModel` などは変更なし）。

**NuGet パッケージ設計:**

```
DotnetPoi.Core     → 全フォーマット実装 (XSSF, HSSF, XWPF, XSLF, POIFS, HWPF, HSLF + SS インターフェース)
DotnetPoi.Formula  → 数式評価器のみ (DotnetPoi.Core に依存)
```

- `DotnetPoi.Formula` のインストールで `createFormulaEvaluator()` が自動有効化。
- `DotnetPoi.Formula` 未インストール時は `NotSupportedException` をスロー。

**XSSFCreationHelper の拡張:**

- `RegisterFormulaEvaluatorFactory()` 静的メソッドでファクトリ登録。
- `createFormulaEvaluator()` は遅延探索機能つき: ランタイムに `DotnetPoi.Formula` アセンブリが存在すれば静的に自動検出。
- `TryAutoRegisterFactory()` が `Type.GetType()` + `RuntimeHelpers.RunClassConstructor()` で発見・起動。

**依存グラフ (最終形):**

```
src/DotnetPoi.Core/    (1 アセンブリ = 全実装)
src/DotnetPoi.Formula/ (Core に依存)
```

- **テスト結果 (全 236 tests 正常):**
- Core.Tests: 195/195
- Formula.Tests: 10/10
- Interop.Tests: 31/31

**移行後のディレクトリ構成:**

```
src/
├── DotnetPoi.Core/       # NuGet: DotnetPoi.Core
│   ├── SS/               # インターフェース・enum・XML writer
│   ├── POIFS/            # OLE2 file system
│   ├── XSSF/             # xlsx format
│   ├── HSSF/             # xls format
│   ├── HWPF/             # doc format
│   ├── XWPF/             # docx format
│   ├── XSLF/             # pptx format
│   └── HSLF/             # ppt format
└── DotnetPoi.Formula/     # NuGet: DotnetPoi.Formula

tests/
├── DotnetPoi.Core.Tests/     # 全フォーマットのテスト (195 tests)
├── DotnetPoi.Formula.Tests/  # 数式評価器のテスト (10 tests)
├── DotnetPoi.Interop.Tests/  # Java POI との相互運用テスト (31 tests)
└── test-files/               # 共有テストデータファイル
```

**全 11 の Example プロジェクト:** 正常ビルド確認済み。

### 移行ガイド (ユーザー向け)

```xml
<!-- 通常の使用 (Core のみ) -->
<PackageReference Include="DotnetPoi.Core" Version="..." />

<!-- 数式評価も使う場合 (Formula を追加) -->
<PackageReference Include="DotnetPoi.Core" Version="..." />
<PackageReference Include="DotnetPoi.Formula" Version="..." />
```

## 2026-05-05 12:xx JST - Formula evaluator split into DotnetPoi.Formula

- Current task: 数式評価器を DotnetPoi.Formula プロジェクトに分離し、DotnetPoi.Core（XSSF/HSSF 等）と独立した NuGet としてビルド可能にした。
- Scope: コミットしない。

### やったこと

**ICell インターフェース拡張（DotnetPoi.SS）:**

- `void setCachedFormulaResult(CellValue value)` を ICell に追加。
- HSSFCell にスタブ実装、XSSFCell に移譲実装（既存の internal SetFormulaCachedValue から呼出）。

**DotnetPoi.Formula 新規プロジェクト作成:**

- `src/DotnetPoi.Formula/DotnetPoi.Formula.csproj` — DotnetPoi.SS のみに依存。
- `FormulaEvaluator` クラス（旧 `XSSFFormulaEvaluator`）を `DotnetPoi.Formula` 名前空間に配置。
- すべての具象 XSSF 型依存（`XSSFWorkbook`, `XSSFCell`, `XSSFSheet`）を SS インターフェース（`IWorkbook`, `ICell`, `ISheet`）に置換。
- `evaluateAll()` はインターフェースメソッド（`getNumberOfSheets()/getSheetAt()`, `getLastRowNum()/getRow()`, `getLastCellNum()/getCell()`）で実装。

**DotnetPoi.XSSF:**

- `DotnetPoi.Formula` へのプロジェクト参照を追加。
- `XSSFFormulaEvaluator.cs` を削除（`DotnetPoi.Formula` に移動）。
- `XSSFCreationHelper.createFormulaEvaluator()` が `FormulaEvaluator` を返すよう更新。

**依存グラフ:**

```
DotnetPoi.SS  ← DotnetPoi.Formula  ← DotnetPoi.XSSF
                (NuGet: DotnetPoi.Formula)
                                        ↓
                                    DotnetPoi.POIFS
```

**テストプロジェクト新規作成:**

- `tests/DotnetPoi.Formula.Tests/DotnetPoi.Formula.Tests.csproj`
  - 参照: `DotnetPoi.Formula` + `DotnetPoi.XSSF`（ワークブック作成に使用）
- `FormulaEvaluatorTests.cs` — 10 tests:
  - SUM/AVERAGE/MIN/MAX/COUNT/CONCATENATE 評価
  - 四則演算、文字列結合、ブール式
  - ゼロ除算エラー、自己参照エラー
  - レンジ参照、`evaluateInCell`、インターフェース経由の生成

### Verification

- `dotnet build src/DotnetPoi.*/` — 全9プロジェクト正常ビルド（0 errors, 0 warnings）。
- `dotnet test tests/DotnetPoi.Formula.Tests/...` — Passed! 10/10.
- `dotnet test tests/DotnetPoi.XSSF.Tests/...` — Passed! 32/32（回帰なし）。

## 2026-05-05 11:xx JST - Phase 7 xlsx fill/border/alignment read

- Current task: xlsx fill/border/alignment 読み取り実装。
- Scope: コミットしない。

### やったこと

`ICellStyle` にアライメントメソッド10個を追加し、`XSSFCellStyle` / `XSSFWorkbook` / `HSSFCellStyle` に実装。

**ReadStyles の拡張（XSSFWorkbook.cs）:**

- `ReadFills()` — `<fills>` セクションをパース。patternType と fgColor indexed を読み取り、`_fills` リストに `XSSFCellStyle` として格納（index 0/1 は組込 default、それ以降がユーザー定義）。
- `ReadBorders()` — `<borders>` セクションをパース。left/right/top/bottom 各辺の style 属性を読み取り、`_borders` リストに格納。
- `ReadCellXfs()` — `fillId` / `borderId` に対応する `_fills[fillId]` / `_borders[borderId]` からプロパティをコピー。`applyFill` / `applyBorder` / `applyAlignment` をパース。`<alignment>` 子要素（horizontal, vertical, wrapText, indent, textRotation）をパース。
- ヘルパーメソッド追加: `ParseHorizontalAlignment()` / `ParseVerticalAlignment()` / `GetHorizontalAlignmentName()` / `GetVerticalAlignmentName()` / `ParseBorderStyleName()`。

**WriteCellXf の拡張:**

- `applyAlignment` 属性と `<alignment>` 子要素を書き出すようになった（horizontal, vertical, wrapText, indent, textRotation）。

**XSSFCellStyle.cs:**

- `_fillRegistered` / `_borderRegistered` フラグで同じ style が `_fills` / `_borders` に重複登録されるのを防止。

**ICellStyle.cs:**

- アライメントメソッド10個を追加: `getAlignment` / `setAlignment` / `getVerticalAlignment` / `setVerticalAlignment` / `getWrapText` / `setWrapText` / `getIndention` / `setIndention` / `getRotation` / `setRotation`。

**VerticalAlignment enum 新規作成:**

- `DotnetPoi.SS.UserModel.VerticalAlignment` を追加。

**HSSFCellStyle.cs:**

- 新しい ICellStyle メソッド10個のスタブ実装を追加（既存のスタブパターンに従う）。

### 追加テスト（XSSFWorkbookTests.cs）

- `RoundTrip_StyledCell_FillRestored` — SolidForeground + Yellow fill の round-trip。
- `RoundTrip_StyledCell_BorderRestored` — 4辺に Medium/Dotted/Thick/Dashed を設定。
- `RoundTrip_StyledCell_AlignmentRestored` — Center/Top/wrapText/indent 1/rotation 45。

### Verification

- `dotnet test tests/DotnetPoi.XSSF.Tests/...` passed (32 tests, 新規3件含む)。
- 全 `dotnet build` source projects 正常。
- agents.md の進捗表を `[~]` → `[x]` に更新（fill/border/alignment read）。

## 2026-05-05 10:xx JST - xlsx style round-trip

- Current task: xlsx round-trip スタイル確認テスト追加。
- Scope: コミットしない。

### 対応状況の整理

`ReadStyles` / `ReadFonts` / `ReadCellXfs` で読み取り済み：
- Font: 名前・bold・italic・strikeout・underline・size・color（indexed）
- DataFormat: numFmtId、カスタム format code

読み取り未実装（fill/border/alignment は `ReadCellXfs` が `fillId`/`borderId`/alignment を読まない）：
→ テストスコープを実装済み範囲に限定し、未対応属性はコメントで明記。

### 追加テスト（XSSFWorkbookTests.cs）

`RoundTrip_StyledCell_FontAndDataFormatRestored`
- Arial 14pt bold italic red + format "0.00" の style を書いて読み返す
- フォント名・高さ・bold・italic・indexed color・format code が全て復元される

`RoundTrip_MultipleStyles_EachCellRestoresItsOwnStyle`
- A1（bold 12pt）と B1（italic + "#,##0.0"）の 2 種スタイルを書いて読み返す
- 各セルが独立して正しいスタイルを保持する

`RoundTrip_BuiltinDateFormat_DataFormatIndexRestored`
- 組み込み日付フォーマット（index 14）を書いて読み返す
- format index が正確に復元される

### Verification

- `dotnet test tests/DotnetPoi.XSSF.Tests/...` passed (29テスト、スタイル round-trip 3件追加)。
- 全スイート異常なし。

## 2026-05-05 09:xx JST - docm/pptm interop + XWPF round-trip

- Current task: docm/pptm Java interop, XWPF round-trip テスト追加。
- Scope boundary: コミットしない。

### docm interop

C#: `Write_DocmWithParagraphsAndVba_CreatesFixtureForPoi`
- `example.docm` の `word/vbaProject.bin` を抽出し `XWPFDocument` に埋め込み
- "from dotnet-poi docm"（bold）と "second paragraph"（italic）の 2 段落
- `phase-docm-interop.docm` に書き出し

Java: `readPhaseDocmInterop`
- content type に "macroEnabled" を含む
- 2 段落：テキスト + bold/italic 属性を確認
- VBA パート（application/vnd.ms-office.vbaProject）が存在

注: POI Java の `XWPFDocument` / `XMLSlideShow` に `isMacroEnabled()` メソッドはない。`getPackagePart().getContentType()` でコンテンツタイプ確認に切り替え。

### pptm interop

C#: `Write_PptmWithSlideAndVba_CreatesFixtureForPoi`
- `example.pptm` の `ppt/vbaProject.bin` を抽出し `XMLSlideShow` に埋め込み
- 画像入り 1 スライド構成
- `phase-pptm-interop.pptm` に書き出し

Java: `readPhasePptmInterop`
- content type に "macroEnabled" を含む
- スライド数 1、shape 1（XSLFPictureShape）
- VBA パート存在

### XWPF round-trip

既存 `Read_WrittenDocument_RestoresTextAndPicture` に bold 属性確認を追加（不足していた）。

新テスト `RoundTrip_MultipleParagraphs_TextAndFormattingRestored`：
- 3 段落（bold/italic/複数ラン）を書いて読み返す
- テキスト、bold、italic、ラン分割が全て復元されることを確認

### ファイル変更

- `tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj`: `example.docm`, `example.pptm` を content item 追加
- `WriteForPoiTests.cs`: `ExtractZipEntry` ヘルパー追加、docm/pptm テスト追加
- `ReadFromDotnetTest.java`: docm/pptm テスト追加、`assertNotNull` import 済み
- `XWPFDocumentTests.cs`: round-trip テスト 1 件追加 + 既存テストに bold 確認追記

### Verification

- C# 全スイート通過（XWPF: 19、Interop: 31、他変化なし）。
- Java `ReadFromDotnetTest` 16テスト通過（+2）。

## 2026-05-05 08:xx JST - Phase 7 assessment, formula evaluator dropped, xlsm interop

- Current task: write Phase 7 current-state assessment into AGENTS.md, drop formula evaluator, implement xlsm interop.
- Scope boundary: AGENTS.md updated (English + Japanese), no commit.

### AGENTS.md updates

- Phase 5 (英日両方): "永久凍結" — `XSSFFormulaEvaluator` は既存テスト用に残すが拡張しない。数式評価はスコープ外。
- Phase 7 step 1〜5: 現在地を `[x]`/`[~]`/`[ ]` で明示。モデル層の残差異（fileVersion等）をノートとして記録。
- Phase 7 進捗表（step別パーセンテージ）を追加。

### xlsm interop 実装

C# 側: `WriteForPoiTests.Write_XlsmWithCellsAndVba_CreatesFixtureForPoi`
- `example.xlsm` から VBA バイトを取得（csproj に content item 追加）
- `XSSFWorkbook` で "MacroSheet" シート + A1="from dotnet-poi xlsm" + B1=99.5 + setVBAProject
- `from-dotnet-poi/phase-xlsm-interop.xlsm` に書き出し

Java 側: `ReadFromDotnetTest.readPhaseXlsmInterop`
- `isMacroEnabled()` == true
- Sheet "MacroSheet" + A1 = "from dotnet-poi xlsm" + B1 = 99.5
- OPC パッケージに vbaProject.bin (application/vnd.ms-office.vbaProject) が存在し非空
- `assertNotNull` import 追加

### Verification

- `dotnet test tests/DotnetPoi.Interop.Tests/cs/...` passed (29 tests, was 28).
- `mvn test -Dtest=ReadFromDotnetTest` passed (14 tests, was 13).
- 既存全テストスイート異常なし。

## 2026-05-05 07:xx JST - Namespace, attribute order, and semantic XSSF tests (items 7-12)

- Current task: implement `XMLBEANS_XML_OUTPUT_TODO.md` Implementation Order items 7-12.
- Scope boundary from user: follow AGENTS.md, update `CHECKPOINT.md` while working, and do not commit.

### Items 7 & 8 — Namespace tests and implementation

New file: `tests/DotnetPoi.SS.Tests/Xml/PoiXmlWriterNamespaceTests.cs` (8 tests)

Tests added:
- Default namespace declaration (`xmlns="..."`) on root element
- `xmlns:r` relationship prefix — both the `WriteAttributeString("xmlns:r", ...)` and prefix overload
- Full spreadsheet workbook root pattern: default ns + `xmlns:r`
- Drawing root pattern: `xmlns:xdr`, `xmlns:a`, `xmlns:r` (in POI order)
- No synthetic `main:` prefix: elements written without prefix don't acquire `main:`
- Prefixed elements use caller-supplied prefix, not a synthetic one
- No duplicate namespace declarations when caller writes once

Implementation (item 8): **no production code changes needed**. `PoiXmlWriter` already passes through namespace declarations as plain attributes and does not sort, hoist, or deduplicate them. The tests serve as the specification.

### Item 9 — Attribute order tests

New file: `tests/DotnetPoi.SS.Tests/Xml/PoiXmlWriterAttributeOrderTests.cs` (5 tests)

Tests added:
- Page margins in POI order (`left`, `right`, `top`, `bottom`, `header`, `footer`) — not alphabetical
- Reverse-alphabetical order (`z`, `a`, `m`) — proves no sorting
- `.rels` relationship `Id`, `Type`, `Target` order
- Two sibling elements each with independent attribute order
- Namespace declarations also follow caller order (xdr before a before r)

Implementation: **no production code changes needed**. The writer already preserves caller attribute order.

### Items 10-12 — Semantic XSSF tests using POI integration fixtures

New file: `tests/DotnetPoi.Interop.Tests/cs/PoiIntegrationFixtureTests.cs` (11 tests)

Tests read from `tests/DotnetPoi.Interop.Tests/fixtures/poi-integration/_workbooks/`:
- `poi-integration-shared-strings-basic.xlsx`: 3 sheets (Sheet1/rich test/Sheet3), A1="Lorem", B1=111.0, A2="ipsum", B2=222.0
- `poi-integration-shared-strings-escaping.xlsx`: first cell contains literal `<` (decoded from `&lt;`)
- `poi-integration-styles-formatting.xlsx`: 3 sheets, Sheet1 A1 = "Dates, all 24th November 2006"
- `poi-integration-comments-write-read.xlsx`: 3 sheets (Sheet1/Sheet2/Sheet3, "AllANumbers"/"AllBStrings" are defined names not sheets), A1="A1", B1="B1", A2=22.3, A3=24.5
- `poi-integration-xlsm-vba-preserve.xlsm`: HasMacros=true, 3 sheets (SheetA/SheetB/SheetC), VBA bytes preserved byte-for-byte on round-trip

Fixture path helper: `GetPoiIntegrationFixturePath` traverses up from `AppContext.BaseDirectory` to find `poi-integration/_workbooks/`. If fixture doesn't exist, test message directs user to run Maven generator.

Item 11: no lexical mismatches found; all semantic tests passed without needing additional `PoiXmlWriter` slices.

Item 12: no fixture-specific XML payloads introduced into `XSSFWorkbook`.

### Verification

- `dotnet test tests/DotnetPoi.SS.Tests/...` passed (91 tests, was 78 before items 7/9).
- `dotnet test tests/DotnetPoi.XSSF.Tests/...` passed (26 tests, unchanged).
- `dotnet test tests/DotnetPoi.XWPF.Tests/...` passed (18 tests, unchanged).
- `dotnet test tests/DotnetPoi.XSLF.Tests/...` passed (25 tests, unchanged).
- `dotnet test tests/DotnetPoi.Interop.Tests/cs/...` passed (28 tests, was 17 before items 10-12).
- All commands still show the existing NU1603 warning for `Microsoft.NET.Test.Sdk`.

## 2026-05-05 06:xx JST - Escaping tests and implementation (items 5 & 6)

- Current task: implement `XMLBEANS_XML_OUTPUT_TODO.md` Implementation Order items 5 and 6.
  - Item 5: Add escaping tests for text and attributes (all chars listed in the TODO).
  - Item 6: Implement only the escaping differences proven to diverge from `System.Xml.XmlWriter`.
- Scope boundary from user: follow AGENTS.md, update `CHECKPOINT.md` while working, and do not commit.

### Evidence used

| Context | Char | POI output | Source |
|---|---|---|---|
| Text | `&` | `&amp;` | `xmlbeans-shared-strings-escaping__poi-options.xml` |
| Text | `<` | `&lt;` | same |
| Text | `>` | literal `>` | same — "A&amp;B &lt;C> \"quoted\" 'single'" |
| Text | `"` | literal `"` | same |
| Text | `'` | literal `'` | same |
| Attribute | `&` | `&amp;` | `poi-integration-hyperlinks__xl__worksheets___rels__sheet1.xml.rels` |
| Attribute | `"` | `&quot;` | `poi-integration-styles-formatting__xl__styles.xml` (formatCode) |
| Attribute | `\` | literal `\` | same (formatCode yyyy\\-mm\\-dd) |

`System.Xml.XmlWriter` escaping (measured with a small C# program):
- Text: `>` → `&gt;`, `"` → literal, `'` → literal
- Attributes: `>` → `&gt;`, `"` → `&quot;`, `'` → literal, `\` → literal

### Proven divergences (POI ≠ SXW)

1. **`>` in text content**: POI = literal, SXW = `&gt;`.

### Bugs in original PoiXmlWriter (diverged from both POI and SXW)

2. **`'` in attribute values**: original code produced `&apos;`; both POI and SXW leave `'` literal in double-quoted attributes.

### Implementation (item 6)

Changed `EscapeCore` in `PoiXmlWriter`:
- Removed `case '>'` entirely — `>` is now literal in both text and attributes.
  - Text: matches POI (proven divergence from SXW fixed).
  - Attributes: consistent with XML spec (double-quoted attributes don't require `>` escaping); no POI fixture contradicts this.
- Removed `case '\'' when forAttribute:` — `'` is now literal in attributes (bug fix; matches both POI and SXW).

### Tests added (item 5)

New file: `tests/DotnetPoi.SS.Tests/Xml/PoiXmlWriterEscapingTests.cs`

Text: `&`, `<`, `>` (literal), `"` (literal), `'` (literal), tab, newline, mixed XMLBeans observation.
Attributes: `&`, `"`, `'` (literal), `\`, relationship URL with `&`, format code with `"`, `<`.

3 tests failed before the fix (`GreaterThanInText`, `ApostropheInAttribute`, `MixedSpecialChars`), all pass after。

### Verification

- `dotnet test tests/DotnetPoi.SS.Tests/...` passed (78 tests)。
- `dotnet test tests/DotnetPoi.XSSF.Tests/...` passed (26 tests)。
- `dotnet test tests/DotnetPoi.XWPF.Tests/...` passed (18 tests)。
- `dotnet test tests/DotnetPoi.XSLF.Tests/...` passed (25 tests)。
- All commands still show the existing NU1603 warning for `Microsoft.NET.Test.Sdk`.

## 2026-05-05 05:xx JST - Empty element serialization tests and implementation (items 3 & 4)

- Current task: implement the `XMLBEANS_XML_OUTPUT_TODO.md` Implementation Order items 3 and 4.
  - Item 3: Add focused failing tests for empty element serialization.
  - Item 4: Implement empty-element behavior in `PoiXmlWriter`; use stream/text interception if needed.
- Scope boundary from user: follow AGENTS.md, update `CHECKPOINT.md` while working, and do not commit。
- Implementation decision:
  - `PoiXmlWriter` is already a custom text-writer-based implementation (not a wrapper around `System.Xml.XmlWriter`).
    It writes `/>` directly in `WriteEndElement()` when the start tag has not yet been closed, producing `<tag/>` with no space before the slash。
  - The "narrow stream/text interception layer" option from the TODO is satisfied by design: the writer uses `TextWriter` directly rather than delegating to `System.Xml.XmlWriter`。
  - No new production code was needed; the implementation was already correct。
- Completed (item 3, initial pass):
  - Added `PoiXmlWriterEmptyElementTests` with root, nested, prefixed, and single-attributed empty-element cases。
- Completed (item 4, strengthened coverage):
  - Extended `PoiXmlWriterEmptyElementTests` with three additional cases drawn from real OOXML patterns:
    - Multi-attributed empty element: `<Relationship Id="..." Type="..." Target="..."/>` (covers `*.rels` patterns)
    - Prefixed + attributed empty element: `<a:picLocks noChangeAspect="1"/>` (covers drawing namespace patterns)
    - Empty string write before `WriteEndElement`: confirms `WriteString("")` does not prevent the `<tag/>` form。
- Verification:
  - `dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --filter PoiXmlWriterEmptyElementTests` passed (7 tests)。
  - `dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj` passed (64 tests)。
  - `dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj` passed (26 tests)。
  - All test commands still show the existing NU1603 package-resolution warning for `Microsoft.NET.Test.Sdk 17.8.2` resolving to `17.9.0`.

## 2026-05-05 04:xx JST - XML writer factory/profile layer

- Current task: implement `XMLBEANS_XML_OUTPUT_TODO.md` Implementation Order 2.
- Scope boundary from user: follow AGENTS.md, keep `CHECKPOINT.md` updated, and do not commit。
- Planned implementation:
  - Add a small factory/profile layer around `PoiXmlWriter` so callers choose XMLBeans spreadsheet-part vs OPC package-part output deliberately。
  - Keep declaration serialization in `PoiXmlWriter`, but avoid hard-coding a single global declaration rule across all OOXML parts。
  - Update XSSF/XWPF/XSLF package writers to create writers through the profile layer instead of directly constructing `PoiXmlWriter`。
  - Add focused xUnit coverage for the factory/profile selection。
- Completed:
  - Added `PoiXmlWriterFactory` with explicit profile creation and OOXML package part classification。
  - Classified `[Content_Types].xml`, `*.rels`, and `docProps/core.xml` as OPC package parts; other XML package entries use the XMLBeans profile。
  - Updated XSSF, XWPF, and XSLF `WriteEntry` helpers to create profiled writers and removed duplicated per-part declaration calls from the writer methods。
  - Left Agile encryption XML on direct `PoiXmlWriter` construction because that XML payload intentionally has no OOXML ZIP part declaration。
- Verification:
  - `dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --filter "PoiXmlWriterFactoryTests|PoiXmlWriterDeclarationProfileTests"` passed。
  - `dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj` passed。
  - `dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj` passed。
  - `dotnet test tests/DotnetPoi.XWPF.Tests/DotnetPoi.XWPF.Tests.csproj` passed。
  - `dotnet test tests/DotnetPoi.XSLF.Tests/DotnetPoi.XSLF.Tests.csproj` passed。
  - All dotnet test commands still show the existing NU1603 package-resolution warning for `Microsoft.NET.Test.Sdk 17.8.2` resolving to `17.9.0`.

## 2026-05-05 03:xx JST - XML declaration profile focused tests

- Current task: implement `XMLBEANS_XML_OUTPUT_TODO.md` Implementation Order 1 by adding focused `PoiXmlWriter` tests for XML declaration profiles。
- Scope boundary from user: keep production implementation minimal for now; do not make fixture-specific `XSSFWorkbook` changes; do not commit。
- Target profiles:
  - XMLBeans spreadsheet parts: `<?xml version="1.0" encoding="UTF-8"?>` followed by a newline, no `standalone`。
  - OPC package parts: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` immediately followed by the root element, no forced newline。
- Implementation approach: add a narrow declaration-profile API on `PoiXmlWriter` only, then characterize both profiles with byte-level tests。
- Completed: added `PoiXmlDeclarationProfile` and focused tests in `PoiXmlWriterDeclarationProfileTests`。
- Verification: `dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --filter PoiXmlWriterDeclarationProfileTests` passed; full `dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj` passed with the existing NU1603 package-resolution warning。

## 2026-05-05 02:xx JST - Agreed recovery plan for XML parity work

- New direction agreed with user:
  1. Remove the `31e9006 parity` work from production/test code for now, including the fixture-specific XSSF writer changes and associated parity tests/fixtures that forced those changes。
  2. Re-study Java POI XML output behavior and XMLBeans behavior at the correct layer. Add reference code/fixture generators where useful, and write a dedicated Markdown TODO/design file describing observed XMLBeans/POI output patterns, open questions, and the implementation order。
  3. Re-implement the behavior incrementally in `PoiXmlWriter`, with focused failing tests per low-level XML divergence。Keep higher-level `XSSFWorkbook` output POI-model-driven, and preserve unknown/original package parts byte-for-byte where the model does not yet support them。
- Important boundary: XML lexical quirks such as declaration format, empty element form, escaping, attribute order, namespace placement, and whitespace belong in `PoiXmlWriter` or focused helpers。Specific workbook content such as defined names, Office revision GUIDs, local absolute paths, workbook extLst contents, and fixture-specific relationship ordering must not be generalized into `XSSFWorkbook` unless directly backed by the POI model/source behavior。
- Do not commit via LLM。
- Step 1 status: completed in working tree by restoring `src/` and the parity-related `tests/` paths back to `7a4b778` (the parent of `31e9006`）。これにより、fixture-specific XSSF writer changes、added XSSF/XWPF/XSLF parity tests、expanded Java parity fixture generator、extra generated xml-parity fixtures が削除される。`dotnet test` は通過する。
- Step 2 direction: do not add XMLBeans as a submodule yet。Add Java probe/fixture generator coverage instead and document the work in `XMLBEANS_XML_OUTPUT_TODO.md`。Use XMLBeans 5.3.0 source jars/tagged source only if executable probes are not enough。
- Step 3 status: added `XMLBEANS_XML_OUTPUT_TODO.md`, `XmlBeansOutputProbeTest.java`, and initial generated XMLBeans probe fixtures under `tests/DotnetPoi.Interop.Tests/fixtures/xmlbeans-output/`。Verified with `mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=XmlBeansOutputProbeTest`。
- Additional fixture strategy: use upstream POI integration-level tests as scenario sources to improve fixture realism。Added `POI_INTEGRATION_FIXTURE_TODO.md` and updated `XMLBEANS_XML_OUTPUT_TODO.md` with a step-by-step plan。Keep POI-derived fixtures semantic-first; do not turn fixture-specific XML into generalized `XSSFWorkbook` output。
- Fixture collection progress: surveyed POI XSSF/model/openxml4j tests and `poi/test-data`。Shortlisted 8 selected candidates in `POI_INTEGRATION_FIXTURE_TODO.md`: shared strings basic, shared strings escaping, styles formatting, comments write/read, pictures multi-sheet, relationships hyperlinks/comments, xlsm VBA preserve, and rich text whitespace preserve。First implementation target is `poi-integration-shared-strings-basic`。
- Fixture collection progress: added `PoiIntegrationFixtureGeneratorTest.java` and generated the first POI-derived integration fixture, `poi-integration-shared-strings-basic`, from POI `sample.xlsx` / `TestSharedStringsTable.testReadWrite` scenario。Output is one xlsx package plus 16 extracted XML/rels files under `tests/DotnetPoi.Interop.Tests/fixtures/poi-integration/`。Verified with `mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=PoiIntegrationFixtureGeneratorTest`。
- Fixture collection progress: extended `PoiIntegrationFixtureGeneratorTest.java` with `poi-integration-shared-strings-escaping` from POI `TestSharedStringsTable.testBug48936` and `poi-integration-styles-formatting` from POI `TestStylesTable.testLoadSaveLoad`。Total generated POI integration fixture files are now 40: three workbook packages plus extracted XML/rels。Verified with the same Maven command。
- Fixture collection progress: extended `PoiIntegrationFixtureGeneratorTest.java` with `poi-integration-comments-write-read`, `poi-integration-pictures-multi-sheet`, and `poi-integration-relationships-hyperlinks-comments`。Total generated POI integration fixture files are now 90: six workbook packages plus extracted XML/rels。Verified with `mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=PoiIntegrationFixtureGeneratorTest`。
- Fixture collection progress: extended `PoiIntegrationFixtureGeneratorTest.java` with `poi-integration-xlsm-vba-preserve` from POI `TestXSSFBugs.bug45431` and `poi-integration-rich-text-space-preserve` from POI `TestXSSFRichTextString` whitespace cases。Extraction now includes `.bin` and `.vml` as well as XML/rels so macro/VML preservation fixtures include `xl/vbaProject.bin` and `xl/drawings/vmlDrawing1.vml`。Total POI integration fixture files are now 123: eight workbook packages plus extracted package entries。Verified with the same Maven command。
- Fixture collection progress: extended `PoiIntegrationFixtureGeneratorTest.java` with `poi-integration-defined-names-print-titles`, `poi-integration-hyperlinks`, `poi-integration-sheet-layout`, and `poi-integration-formula-recalculation`。These cover defined names/print titles, external hyperlink relationships and query escaping, row/column layout with panes/grouping/merged regions, and workbook formula recalculation XML。Total POI integration fixture files are now 165: twelve workbook packages plus extracted package entries。Verified with `mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=PoiIntegrationFixtureGeneratorTest`。
- Fixture analysis progress: updated `XMLBEANS_XML_OUTPUT_TODO.md` with observations from the XMLBeans probes and 12 POI integration fixture cases。Key split: spreadsheet XMLBeans parts use UTF-8 declaration plus newline and no `standalone`, while OPC package-level XML parts such as `[Content_Types].xml`, `.rels`, and core properties use `standalone="yes"` with no forced newline before the root。Recorded `PoiXmlWriter` implementation order: declaration profiles, `<tag/>` empty elements, escaping, namespace behavior, attribute-order preservation, then semantic fixture tests and unknown-part preservation。

## 2026-05-05 02:xx JST - Review of recent XSSF writer/parity work

- User raised concern that recent commits may have made XSSF workbook XML writing ad hoc and destabilizing。
- Reviewed last 5 commits。Main risky commit is `31e9006 parity`; `3830f84` adjusts generated fixtures/timestamps。
- `dotnet test` passes locally, but `XSSFWorkbook.WriteWorkbook` now contains fixture-specific-looking constants: Office revision namespaces, GUIDs, a local absolute `x15ac:absPath`, hard-coded defined names, workbook window values, and extension list data。
- Assessment: the concern is valid。The tests currently prove parity for a narrow fixture, not a general POI-faithful writer。Recommended next step is to quarantine this behavior behind preservation of original package parts or fixture-only tests, then revert generalized writer output to minimal/POI-derived data。

## 2026-05-05 01:54 JST - XML parity CI drift

- GitHub Actions `Verify XML Parity Fixtures` failed because `XmlParityFixtureGeneratorTest` rewrites `xlsm-basic` by opening `tests/test-files/example.xlsm` with Apache POI and saving it through `XSSFWorkbook.write()`。
- The committed `xlsm-basic__*.xml` fixtures are intentionally hybrid: DotnetPoi regenerates workbook/content-types/relationships but preserves unchanged macro workbook parts such as doc props, styles, shared strings, worksheets, drawings, theme, and calcChain。
- Fix: updated `generateMacroEnabledXlsm` to first generate the POI package, then overlay the preserved xlsm entries from the source workbook so CI regenerates the same hybrid fixture set。
- Verification: `dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --filter XmlParity_XlsmBasic_MatchesPoiFixtures` passes。Local Maven is not installed, so the exact GitHub workflow command still needs CI or a Maven-equipped environment。

## 2026-05-10 11:xx JST - Round-trip batch 1: merge cells + column width / row height

- **Merge cells**: Implemented read + write + round-trip test。
  - Added `CellRangeAddress` class in `DotnetPoi.SS.Util` with `Parse` and `FormatAsString`。
  - Added `addMergedRegion()` / `getMergedRegions()` to `ISheet` interface。
  - Implemented in `XSSFSheet` (store list, write `<mergeCells>`/`<mergeCell>` in `WriteWorksheet`, parse in `ReadWorksheet`)。
  - Added stubs in `HSSFSheet` for interface completeness。
  - Test: `RoundTrip_MergeCells_Preserved` — writes 2 merged regions, reads back, verifies all coordinates。
- **Column width**: Implemented read + write + round-trip test。
  - Added `setColumnWidth()` / `getColumnWidth()` to `ISheet` interface。
  - Implemented in `XSSFSheet` (dictionary, write `<cols>`/`<col>` in `WriteCols`, parse in `ReadWorksheet`)。
  - Test: `RoundTrip_ColumnWidth_Preserved` — sets 80-char and 40-char widths, verifies approximate round-trip。
- **Row height**: Implemented read + write + round-trip test。
  - Added `setHeight()` / `getHeight()` to `IRow` interface。
  - Implemented in `XSSFRow` (store float, write `ht`/`customHeight` attributes in `WriteRow`, parse in `ReadWorksheet`)。
  - Test: `RoundTrip_RowHeight_Preserved` — sets 45 pt custom height, verifies round-trip; default row returns 15 pt。
- All 191 Core tests pass (188 existing + 3 new)。Formula (10) and Interop (31) も pass。

## 2026-05-10 JST - Round-trip batch 2: hyperlinks

- **Hyperlinks**: Implemented read + write + round-trip test。
  - Created `IHyperlink` interface and `HyperlinkType` enum in `DotnetPoi.SS.UserModel`。
  - Created `XSSFHyperlink` class with `Address`, `CellRef`, `Type`, `RelationshipId`, `IsExternal` properties。
  - Added `getHyperlink()`/`setHyperlink()` to `ICell` interface。
  - Implemented in `XSSFCell` (store `_hyperlink` field, auto-register via `sheet.AddHyperlink` on set)。
  - Added stubs in HSSFCell for interface completeness。
  - Write: `WriteHyperlinks()` emits `<hyperlinks>` section with `<hyperlink ref="..." r:id="..."/>`。
  - Write: `WriteSheetRelationships()` emits rels entries with `TargetMode="External"` for URLs。
  - Write: `AssignHyperlinkRelationshipIds()` pre-assigns relationship IDs before writing。
  - Write: Sheet rels file is now written when hyperlinks exist (not only when drawing exists)。
  - Read: `ReadSheetHyperlinkRelationships()` reads sheet rels, builds dictionary of relId → (URL, isExternal)。
  - Read: `<hyperlink>` elements parsed in `ReadWorksheet`, hyperlinks created and associated with existing cells。
  - Added `ParseCellRef()` helper for parsing "A1"-style references。
  - Fixed `XSSFHyperlink` constructor to set `IsExternal = true` for URL/File/Email types。
  - Test: `RoundTrip_Hyperlink_Preserved` — creates URL hyperlink, writes, reads back, verifies address, cellRef, type。
- All 192 Core tests pass (191 + 1 new)。Formula (10) and Interop (31) も pass。

## 2026-05-10 JST - Round-trip batch 3: print settings (page layout, header/footer)

- **Page settings**: Implemented read + write + round-trip test。
  - Added page margin properties to `XSSFSheet` (PageMarginBottom/Footer/Header/Left/Right/Top, double, default OOXML values)。
  - Added page setup properties (PageOrientation, PaperSize, Scale, FitToWidth, FitToHeight)。
  - Added header/footer properties (HeaderCenter, FooterCenter)。
  - Write: Page setup `<pageSetup>` element written only when non-default properties set。
  - Write: Page margins `<pageMargins>` written from sheet properties (no longer hardcoded)。
  - Write: Header/footer `<headerFooter>` with `<oddHeader>`/`<oddFooter>` children。
  - Read: Parse `<pageSetup>`, `<pageMargins>`, `<oddHeader>`, `<oddFooter>` in ReadWorksheet。
  - Added `ParseDoubleAttr()` helper for attribute reading with default fallback。
  - Test: `RoundTrip_PrintSettings_Preserved` — sets orientation, paper size, margins, header/footer; writes and reads back。
- All 193 Core tests pass (192 + 1 new)。Formula (10) and Interop (31) も pass。

## 2026-05-10 JST - Round-trip batch 4: data validation

- **Data validation**: Implemented read + write + round-trip test。
  - Created `DataValidationType` and `DataValidationOperator` enums in `SS/UserModel`。
  - Created `XSSFDataValidation` class with properties: Sqref, Type, Operator, Formula1/2, AllowBlank, ShowInputMessage, ShowErrorMessage, ShowDropDown, ErrorStyle, ErrorTitle/Message, PromptTitle/Message。
  - Write: `<dataValidations>` element with `<dataValidation>` children, written via `WriteDataValidations()`。
    - Writes `type`, `operator`, `allowBlank`, `showInputMessage`, `showErrorMessage`, `showDropDown`, `errorStyle`, `errorTitle`, `error`, `promptTitle`, `prompt`, `sqref` attributes。
    - Writes `<formula1>` / `<formula2>` child elements。
  - Read: Parses `<dataValidation>` elements in ReadWorksheet, including child formula elements。
  - Added `GetDataValidationTypeName()` and `GetDataValidationOperatorName()` for write path。
  - Added `DataValidationTypeFromName()` and `DataValidationOperatorFromName()` for read path。
  - Test: `RoundTrip_DataValidation_Preserved` — creates whole-number validation (1-100), writes, reads back, verifies all properties。
- All 194 Core tests pass (193 + 1 new)。Formula (10) and Interop (31) も pass。

## 2026-05-10 JST - Round-trip batch 5: conditional formatting

- **Conditional formatting**: Implemented read + write + round-trip test。
  - Created `ConditionalFormatType` enum and `XSSFCFRule` / `XSSFConditionalFormatting` classes in `XSSF/UserModel`。
  - Write: `<conditionalFormatting>` elements with `<cfRule>` children, written via `WriteConditionalFormatting()`。
    - Writes `type`, `priority`, `operator`, `text`, `dxfId` attributes。
    - Writes `<formula>` child elements。
  - Read: Parses `<conditionalFormatting>` and nested `<cfRule>` elements in ReadWorksheet, including child formula elements。
  - Added `GetCfTypeName()` for write path and `CfTypeFromName()` for read path, supporting all OOXML rule types (cellIs, expression, top10, uniqueValues, duplicateValues, containsText, beginsWith, endsWith, containsBlanks, containsErrors, timePeriod, aboveAverage)。
  - Test: `RoundTrip_ConditionalFormatting_Preserved` — creates cellIs > 100 rule, writes, reads back, verifies type/priority/operator/formula。
- All 195 Core tests pass (194 + 1 new)。Formula (10) and Interop (31) も pass。

## 2026-05-10 JST - Round-trip summary (all features complete except Pivot tables / Charts)

All requested round-trip features are now implemented。Pivot tables and Charts are deferred (major features requiring separate XML parts)。

### Implemented features (read + write + round-trip test):

| Feature | Status | Test |
|---------|--------|------|
| Pictures (rotation, image data) | ✅ Already done (XSSFPictureTests) | `Rotation_RoundTrip_PreservesRotation`, `Read_WorkbookWithJpegMedia_RestoresPictureData` |
| Fill / Border / Alignment | ✅ Phase 7 | `RoundTrip_StyledCell_FillRestored`, `BorderRestored`, `AlignmentRestored` |
| Merge cells | ✅ Batch 1 | `RoundTrip_MergeCells_Preserved` |
| Column width / Row height | ✅ Batch 1 | `RoundTrip_ColumnWidth_Preserved`, `RoundTrip_RowHeight_Preserved` |
| Hyperlinks | ✅ Batch 2 | `RoundTrip_Hyperlink_Preserved` |
| Print settings (margins, orientation, header/footer) | ✅ Batch 3 | `RoundTrip_PrintSettings_Preserved` |
| Data validation | ✅ Batch 4 | `RoundTrip_DataValidation_Preserved` |
| Conditional formatting | ✅ Batch 5 | `RoundTrip_ConditionalFormatting_Preserved` |
| Freeze panes | ✅ Batch 6 | `RoundTrip_FreezePane_Preserved` |
| Hidden rows / Hidden columns | ✅ Batch 6 | `RoundTrip_HiddenRow_Preserved`, `RoundTrip_HiddenColumn_Preserved` |
| Shared strings (plain text) | ✅ Batch 6 | `RoundTrip_SharedStrings_Preserved` |

### Total tests: 240 all passing (199 Core + 31 Interop + 10 Formula)

### Remaining deferred features:
- **Rich text formatting** (preserving `<r>` / `<rPr>` runs in shared strings on write-back)
- **Pivot tables** (requires separate `xl/pivotTables/` parts)
- **Charts** (requires separate `xl/charts/` parts)
- **Comments** (requires `xl/comments/` + vmlDrawing parts)
- **XLS (HSSF) encryption** (RC4/XOR different scheme from OOXML Agile)

## 2026-05-10 JST - docx (XWPF) round-trip

| # | Feature | Read | Write | Test | Status |
|---|---------|------|-------|------|--------|
| 1 | Paragraphs & runs (text) | ✅ | ✅ | ✅ | ✅ |
| 2 | Bold / Italic | ✅ | ✅ | ✅ | ✅ |
| 3 | Inline pictures + rotation | ✅ | ✅ | ✅ | ✅ |
| 4 | Run font name / size / color / underline / strikethrough | ✅ | ✅ | ✅ | ✅ |
| 5 | Paragraph alignment (left/center/right/justify) | ✅ | ✅ | ✅ | ✅ |
| 6 | Paragraph indentation / spacing | ✅ | ✅ | ✅ | ✅ |
| 7 | Numbering / bullet lists | ✅ | ✅ | ✅ | ✅ |
| 8 | Tables | ✅ | ✅ | ✅ | ✅ |
| 9 | Hyperlinks | ✅ | ✅ | ✅ | ✅ |
| 10 | Headers / Footers | ✅ | ✅ | ✅ | ✅ |
| 11 | Page setup (size / orientation / margins) | ✅ | ✅ | ✅ | ✅ |

### docx numbering infrastructure (re-added after git checkout revert)

Numbering write/read was lost during a git checkout revert and has been re-implemented:

- **Write**: `WriteNumbering()` method writes `<w:numbering>` with `<w:abstractNum>` (+ lvl/start/numFmt/lvlText) and `<w:num>` (+ abstractNumId) elements.
- **Read**: `ReadNumbering()` method parses `word/numbering.xml` back into `_abstractNums` and `_numInstances` lists, with ID tracking.
- **Content type**: Conditional override for `/word/numbering.xml` in `WriteContentTypes`.
- **Relationship**: Conditional relationship entry in `WriteDocumentRelationships` (dynamic rId calculation after images and vbaProject).
- **WriteParagraph**: Updated to write full pPr (jc, ind, spacing, numPr/numId/ilvl).
- **WriteRun**: Updated to write full rPr (b, i, u, strike, rFonts, sz, color).
- **ReadDocument**: pPr/rPr element handlers for ind, spacing, numPr, numId, ilvl; rFonts, sz, color, underline, strike.
- All 19 XWPF tests pass (12 existing + 3 font/alignment + 2 indent/spacing + 2 bullet/numbered list).

---

## 2026-05-10 JST - PPTX (XSLF) round-trip: text boxes + slide size

### Text box (p:sp) write/read

- **XSLFAutoShape**, **XSLFTextParagraph**, **XSLFTextRun**: New classes for text box model with formatting (bold, italic, underline, strikethrough, font size, font name, color).
- **XSLFSlide**: Added `createTextBox()`, `getAutoShapes()`, `_autoShapes` list.
- **WriteAutoShape**: Writes `p:sp` element with `p:nvSpPr`, `p:spPr` (xfrm with rot/flip/off/ext, prstGeom rect), `p:txBody` (a:bodyPr, a:lstStyle, a:p → a:r → a:t with rPr formatting).
- **ParseSlideXml**: Extended to parse `p:sp` elements: shape ID, anchor, txBody → a:p → a:r → a:t text content, plus a:rPr for all formatting properties.
- **Bug fix**: `WriteStartElement(a,"b")`/`WriteEndElement()` were on same line without braces — `WriteEndElement()` always executed (even when condition false), causing `InvalidOperationException: No open element to close`. Fixed by wrapping in `{ }`.
- **6 round-trip tests**: single-run text, multi-run (bold formatting), multiple paragraphs, anchor preservation, slide size, combined text box + picture.

### Slide size

- **Fields**: `_slideCx`/`_slideCy` (default 9,144,000 × 6,858,000 EMU).
- **Public API**: `getSlideCx()`, `getSlideCy()`, `setSlideSize()`.
- **Write**: `WritePresentation()` uses instance fields instead of hardcoded defaults.
- **Read**: `ParseSlideSize()` reads `p:sldSz` cx/cy from presentation.xml on load.

### Test count

- Core: **219** (212 existing + 5 docx round-trip: tables, hyperlinks, page setup, header/footer + 2 pptx round-trip: unknown parts, table)
- Formula: **10**
- Interop: **31**
- **Total: 260**

### docx completion (tables, hyperlinks, page setup, headers/footers)

Items 8–11 on the docx round-trip checklist are now implemented:

- **Tables (item 8)**: Existing write (`WriteTable`) and read (`ReadDocument` table parsing) paths verified with round-trip tests. 2 test methods added.
- **Hyperlinks (item 9)**: New `XWPFRun.setHyperlink(url)`/`getHyperlink()` API. Write wraps runs in `<w:hyperlink r:id="...">`. Read parses hyperlink relationships and element references. `CollectHyperlinks()` recursively scans top-level and table-cell paragraphs. `BuildHyperlinkRelMap()` reads doc rels. 1 test method.
- **Page setup (item 11)**: New `setPageSize()`, `setLandscape()`, `setMargins()` API. Write emits `<w:pgSz>` (w, h, orient) and `<w:pgMar>` (top, right, bottom, left, header, footer) in sectPr. Read parses pgSz/pgMar elements. 1 test method.
- **Headers/Footers (item 10)**: New `setHeaderText()`/`setFooterText()` API. Write emits `word/header1.xml` and `word/footer1.xml` as separate ZIP entries with content type overrides, relationships, and `<w:headerReference>`/`<w:footerReference>` in sectPr. Read parses rels and header/footer XML to extract text. 1 test method.

---

## 2026-05-10 JST - PPTX (XSLF) round-trip batch 1: unknown part preservation + tables

### Plan

| # | Feature | Priority | Notes |
|---|---------|----------|-------|
| 1 | Unknown part preservation | 🔥 High | ✅ Done — Save ZIP entries not understood by model byte-for-byte |
| 2 | Layouts / Masters / Themes | 🔥 High | ✅ Preserved via unknown parts mechanism (stubs on create, real from file on load) |
| 3 | Tables (p:graphicData > a:tbl) | ⭐ Medium | ✅ Done |
| 4 | Charts | 🕐 Later | Requires separate chart parts + relationships |
| 5 | Notes, Connectors, Group Shapes | 🕐 Later | Extend shape model coverage |

### Batch 1: Unknown part preservation

The `write()` method currently creates a fresh ZIP without copying over parts it doesn't understand (layouts, masters, themes, media, etc.). This means round-tripping a real pptx file (e.g. from PowerPoint or Java POI) loses all design.

**Approach**: During `Load()`, collect all non-model ZIP entries and their raw bytes. During `write()`, emit those entries first, then let the model overwrite the entries it knows about (`ppt/presentation.xml`, `ppt/slides/slide1.xml`, `ppt/slides/_rels/slide1.xml.rels`, etc.).

**Implementation**:
- Added `_preservedEntries` field (`Dictionary<string, byte[]>` with `OrdinalIgnoreCase`).
- `CollectPreservedEntries(ZipArchive)` iterates ZIP entries, excludes known model paths + `ppt/media/*`, and stores the rest.
- `GetModelEntryNames()` returns a `HashSet` of expected paths (presentation, slides, slide rels, slide masters, layouts, theme, content types, rels, core props, vba, image media).
- `Load()` calls `CollectPreservedEntries(archive)` after slides are loaded.
- `write()` emits preserved entries first via `WriteBinaryEntry`, then model entries overwrite.
- Test injects extra `ppt/slideLayouts/layout2.xml` + rels into a pptx, loads, writes, and verifies both entries survive.

### Batch 2: Tables (p:graphicData > a:tbl)

**Write path**:
- `WriteTableGraphicFrame()` writes `p:graphicFrame` element with `nvGraphicFramePr`, `xfrm` (anchor), `a:graphic > a:graphicData` (table URI), `a:tbl` with `tblPr`, `tblGrid` (`gridCol` widths), `a:tr` (`a:tc` with `txBody` containing paragraphs/runs with full formatting support).

**Read path**:
- `ParseSlideXml()` extended with `inGraphicFrame`, `inTbl`, `tableInAP`, `tableInAR`, `tableInARPr` state flags.
- Parses `p:graphicFrame`, extracts `cNvPr` (id), `a:off`/`a:ext` (anchor), `a:tbl`, `a:gridCol` (widths), `a:tr`/`a:tc`, and cell text via `a:txBody > a:p > a:r > a:t` with full `a:rPr` formatting.
- End element handlers properly assemble `XSLFTable → XSLFTableRow → XSLFTableCell → XSLFTextParagraph → XSLFTextRun`.

**Model classes** (new file `XSLFTable.cs`):
- `XSLFTable`: ShapeId, GridColWidths, Rows, setAnchor/createRow.
- `XSLFTableRow`: Cells, createCell.
- `XSLFTableCell`: Paragraphs, addParagraph (reuses XSLFTextParagraph for text content).

**1 round-trip test**: `RoundTrip_Table_Restored` — 2×2 table with cell text, grid widths, anchor position.

## 2026-05-05 JST - GitHub Actions ci.yml syntax fix

- Investigated `.github/workflows/ci.yml` GitHub syntax error reported at line 39.
- Root cause: the unquoted step `name` contained `all formats: xlsx...`; YAML treats `: ` inside an unquoted scalar as mapping syntax.
- Fixed by quoting the affected step name.
- Scope intentionally limited to workflow syntax; did not touch existing unrelated modified files.

## 2026-05-05 JST - EdgeCaseProbeExample formula evaluator reference

- Investigated failing `dotnet run --project examples/EdgeCaseProbeExample/` probe:
  `xlsx formula edge results: divide by zero, circular reference, missing cells`.
- Root cause: the example calls `workbook.getCreationHelper().createFormulaEvaluator()` but referenced only `DotnetPoi.Core`; by design formula evaluation lives in `DotnetPoi.Formula`.
- Added `DotnetPoi.Formula` project reference to `examples/EdgeCaseProbeExample/EdgeCaseProbeExample.csproj`.

## 2026-05-11 JST - xlsx rich text (per-character formatting in shared strings)

**Problem**: When creating a cell with mixed formatting (e.g. bold "Hello " + italic "World"), the formatting was lost on round-trip because the shared strings table only stored/retrieved plain strings.

**Solution**: Full rich text support following Apache POI's `XSSFRichTextString` model.

### Implementation

- **New file `XSSFRichTextString.cs`**: Model class with `TextRun` inner class carrying per-run formatting (Bold, Italic, Underline, Strikethrough, FontSize in points, FontName, Color as hex RGB). Provides `IsRichText` property, `getString()` for plain text concatenation, and `addRun()` overloads.

- **`XSSFCell.cs`**: Added `_richTextStringValue` field, `setCellValue(XSSFRichTextString)`, `getRichStringCellValue()`, and `SetRichTextStringFromSst()` for loading.

- **`XSSFWorkbook.cs`**:
  - `_sharedStrings` changed from `List<string>` to `List<XSSFRichTextString>`.
  - `ReadSharedStrings` now returns `List<XSSFRichTextString>`, parsing `<r>` (text runs) with `<rPr>` (formatting: b, i, u, strike, sz, rFont, color). Falls back to plain `<t>` for legacy SST.
  - `WriteSharedStrings` writes `<r>` elements with `<rPr>` when `IsRichText` is true; writes plain `<si><t>` for plain strings.
  - `BuildSharedStrings` collects `XSSFRichTextString` from cells via `getRichStringCellValue()`.
  - `ApplyCellValue` passes `XSSFRichTextString` to the cell via `SetRichTextStringFromSst()`.
  - `ReadWorksheet` and `ApplyCellValue` signatures updated to `IReadOnlyList<XSSFRichTextString>`.

### Test

- `RoundTrip_RichTextFormatting_Preserved`: Creates a cell with two formatted runs (bold red "Hello " + italic blue "World"), writes to stream, loads back, and verifies plain text, run count, bold/italic flags, font size, font name, and color on each run.

### Test count

- **Core**: 221 (220 existing + 1 rich text)
- **Formula**: 10
- **Interop**: 31
- **Total**: 262

## 2026-05-06 JST - Phase 9 docs site implementation direction

- User requested Phase 9 setup direction first: decide directory structure, language/tooling, and solution approach, then record it in `agents.md`.
- Added Phase 9 implementation policy to both English and Japanese sections of `agents.md`.
- Decision: keep docs generation maintainable by using a small C#/.NET console generator instead of a separate heavy docs framework or another language runtime.
- Generator location: `tools/DotnetPoi.DocsGenerator/`.
- Solution policy: add the generator project to the existing `DotnetPOI.sln`; do not create a separate solution unless the docs tool becomes independently releasable.
- Canonical docs source: `docs_src/site.json`, `docs_src/content/`, `docs_src/assets/`, and `docs_src/templates/`.
- Generated output: `docs/` only; generated HTML should not be edited by hand.
- Build command recorded: `dotnet run --project tools/DotnetPoi.DocsGenerator -- docs_src docs`.
- Next TODO: scaffold `docs_src/`, `docs/`, and `tools/DotnetPoi.DocsGenerator/`; add the project to `DotnetPOI.sln`; implement minimal Markdown-to-HTML generation; add a first getting-started page linked to an existing runnable example.

## 2026-05-06 JST - Phase 9 usage samples started

- User clarified Phase 9 should start by reading the POI documentation and creating real runnable usage samples; class structure/API documentation is not needed.
- Read local Apache POI documentation sources:
  - `poi/src/documentation/content/xdocs/components/spreadsheet/quick-guide.xml`
  - `poi/src/documentation/content/xdocs/components/document/quick-guide-xwpf.xml`
  - `poi/src/documentation/content/xdocs/components/slideshow/quick-guide.xml`
- Added `agents.md` rule: Phase 9 is usage-docs-only for now; do not create API reference, class hierarchy, or architecture pages unless explicitly requested.
- Added `examples/UsageSamples/` console project and included it in `DotnetPOI.sln`.
- `UsageSamples` creates and verifies:
  - `examples/output/usage-workbook.xlsx` — xlsx cells, styles, merged title, freeze pane, data validation, conditional formatting, rich text, readback verification.
  - `examples/output/usage-document.docx` — docx paragraphs, formatted runs, hyperlink metadata, table, image, readback verification.
  - `examples/output/usage-presentation.pptx` — pptx slides, text boxes, image, table, readback verification.
- Added `docs_src/site.json` and `docs_src/content/getting-started/usage-samples.md` as initial usage-only docs source.
- Updated `examples/README.md` with the new runnable sample command.
- Verification passed: `dotnet run --project examples/UsageSamples/UsageSamples.csproj`.
- Next TODO: add focused usage samples/docs for common spreadsheet workflows from POI quick guide: read existing workbook, formulas as text/cached values, images, page setup/print settings, hyperlinks, and macro-enabled round-trip preservation.

## 2026-05-06 JST - Phase 9 local docs generator

- User requested a tool under `tools/` that generates the GitHub Pages-ready `docs/` site from `docs_src/`, with language/framework left to the agent.
- Implemented `tools/DotnetPoi.DocsGenerator/` as a small C#/.NET console app with no external packages.
- Added the generator project to `DotnetPOI.sln`.
- Generator behavior:
  - Reads `docs_src/site.json` for title, description, and navigation.
  - Reads Markdown pages under `docs_src/content/`.
  - Supports the Markdown subset currently needed by usage docs: headings, paragraphs, bullet lists, fenced code blocks, inline code, and links.
  - Copies `docs_src/assets/` to `docs/assets/`.
  - Writes `docs/index.html` plus per-page HTML, using relative links suitable for GitHub Pages project sites.
- Added `docs_src/assets/site.css` and generated:
  - `docs/index.html`
  - `docs/getting-started/usage-samples.html`
  - `docs/assets/site.css`
- Verification passed: `dotnet run --project tools/DotnetPoi.DocsGenerator -- docs_src docs`.
- Removed generated `tools/DotnetPoi.DocsGenerator/bin` and `tools/DotnetPoi.DocsGenerator/obj` after verification.
- Next TODO: as more Markdown pages are added to `docs_src/content/`, add them to `docs_src/site.json` navigation and rerun the generator before pushing `docs/`.

---

## 2026-05-14 JST — xlsx rich text + pivot table programmatic creation + docx unknown part preservation

### docx unknown part preservation (item 1/3)

Added `_preservedEntries` mechanism to `XWPFDocument` (matching existing xlsx/pptx pattern):

- `_preservedEntries` field (`Dictionary<string, byte[]>` with `OrdinalIgnoreCase`)
- `GetModelEntryNames()` returns known model paths (document, settings, rels, content types, numbering, header/footer, images, vba)
- `CollectPreservedEntries(ZipArchive)` stores non-model ZIP entries byte-for-byte
- `Load()` calls `CollectPreservedEntries` after other parsing
- `write()` emits preserved entries first, model entries overwrite
- 1 test: injects `word/styles.xml` and `docProps/custom.xml`, verifies survival through round-trip

### xlsx rich text / shared strings formatting (item 2/3)

New `XSSFRichTextString` class with `TextRun` sub-class supporting per-run formatting (bold, italic, underline, strikethrough, font size, font name, color):

- **Model**: `XSSFRichTextString` stores list of `TextRun` objects; `getString()` concatenates; `IsRichText` property detects formatting.
- **SST storage**: Changed `_sharedStrings` from `List<string>` to `List<XSSFRichTextString>`, keyed by plain text for dedup.
- **Read path**: `ReadRichSi()` parses `<si>` with `<r>` runs + `<rPr>` formatting (b, i, u, strike, sz/100, rFont/latin, srgbClr).
- **Write path**: `WriteSharedStrings()` emits `<r>` elements with full `<rPr>` for formatted runs, plain `<t>` for unformatted.
- **Cell API**: `XSSFCell.setCellValue(XSSFRichTextString)`, `getRichStringCellValue()`, `SetRichTextStringFromSst()`.
- **1 round-trip test**: "Hello " (bold, red, 14pt, Arial) + "World" (italic, blue, 12pt, Calibri).

### xlsx pivot table programmatic creation (item 3/3)

New model classes with write-only support (existing files round-trip via unknown parts):

- **`XSSFPivotTable`**: PivotTableIndex, CacheId, CacheDefinition/CacheRecords/Cache objects, WritePivotTableDefinition().
- **`XSSFPivotCache`**: Simple CacheId container.
- **`XSSFPivotCacheDefinition`**: CacheId, SourceSheetName, SourceRef.
- **`XSSFPivotCacheRecords`**: CacheId container.
- **`XSSFSheet.createPivotTable()`**: Creates pivot table with cache allocation, registers with workbook.
- **`XSSFWorkbook` wiring**:
  - WriteWorkbook: `<pivotCaches>` element with cache references.
  - WriteWorkbookRelationships: pivot cache def relationships with sequential rId.
  - WriteContentTypes: overrides for pivotTable, pivotCacheDefinition, pivotCacheRecords.
  - WriteSheetRelationships: sheet-level pivot table relationships.
  - GetModelEntryNames: includes pivot parts.
  - write(): emits pivotTable{index}.xml, pivotCacheDefinition{id+1}.xml, pivotCacheRecords{id+1}.xml.

### docx fields (TOC / page numbers / mail merge)

New `XWPFField` model class with write and read round-trip support:

- **`XWPFField`**: Model class with `Instruction` (e.g. `" PAGE "`) and `Result` (e.g. `"1"`).
- **`XWPFParagraph.addField(string instruction, string result)`, `getFields()`**: Public API to add and retrieve fields per paragraph.
- **Write path (`WriteParagraph`)**: After runs, emits the full OOXML field sequence: `fldChar begin → instrText → fldChar separate → t (result) → fldChar end`.
- **Read path (`ReadDocument`)**: Parses `fldChar` (begin/separate/end) and `instrText` elements, accumulates field result from `t` after separate, constructs `XWPFField` on end.
- **1 round-trip test**: Creates paragraph with text run + PAGE field + text run, writes, reloads, verifies field instruction (`" PAGE "`) and result (`"1"`).

### Test count

- Core: **226** (+1 new: xlsx autofilter)
- Formula: **10**
- Interop: **51** (+16 preservation fixture tests for unsupported features: charts, comments, textboxes, OLE, SmartArt, video, audio, shapes, footnotes, change tracking, columns, etc.)
- **Total: 287** (Core 226, Formula 10, Interop C# 51)

---

## 2026-05-15 JST — Interop test 整備

### Plan

1. ✅ **docx Interop B (dotnet → Java POI)**: Expanded fixture to include tables, hyperlinks, headers/footers, numbering, page setup, rich text — C# fixture generator + Java POI validator implemented.
2. ✅ **pptx Interop B (dotnet → Java POI)**: Added fixture with text boxes (bold/italic text + font size) — C# fixture generator + Java POI validator implemented.
3. ✅ **docx/pptx Interop A (Java POI → dotnet)**: Java fixture generator + C# validator.

### Implementation

#### docx Interop A

- **Java `WriteForDotnetTest.writePhaseDocxComprehensive()`**: Creates docx with plain paragraph, bold+normal paragraph, italic paragraph, 2×2 table (A1/B1/A2/B2), hyperlink paragraph, centered header ("Interop Header"), centered footer ("Interop Footer").
- **C# `ReadPoiGeneratedTests.Read_DocxComprehensive_GeneratedByPoi()`**: Reads fixture, verifies paragraphs (plain, bold/run split, italic), table structure and cell text, header/footer text.
- Fixture: `fixtures/from-poi/phase-docx-comprehensive.docx`

#### pptx Interop A

- **Java `WriteForDotnetTest.writePhasePptxComprehensive()`**: Creates pptx with text box: 2 paragraphs — "Bold Title" (bold, 18pt) and "Italic subtitle" (italic, 14pt). (Table skipped — POI 5.5.1 `addColumn()` bug on empty table.)
- **C# `ReadPoiGeneratedTests.Read_PptxWithTextBoxes_GeneratedByPoi()`**: Reads fixture, verifies shape type, paragraph count (3 including POI default), and text content.
- Fixture: `fixtures/from-poi/phase-pptx-comprehensive.pptx`
- Note: POI 5.5.1's `XSLFTextRun.isBold()` and `CTTextCharacterProperties.isSetB()` return false for loaded PPTX even when `<a:b/>` is present. Formatting verification relies on C# round-trip tests.

---

## 2026-05-15 JST — 優先順位決定（TODO frozen）

### 決定事項

グラフ（チャート）は実装コストが極めて大きい（30種のチャート型、DrawingML スキーマ、多ファイル構成）ため後回し。代わりに以下を順次実装する。

| 順位 | 項目 | フォーマット | 目安 |
|---|---|---|---|
| 1 | **セル保護 / ブック保護** | xlsx | 半日〜1日 |
| 2 | **フィールド（TOC/ページ番号/差し込み印刷）** | docx | 2〜3日 |
| 3 | **オートフィルター** | xlsx | 1〜2日 |
| 🔵 | グラフ作成 | xlsx, pptx | 後回し |

### 備考

- 数式評価（DotnetPoi.Formula）は「テンプレート填充 → 保存 → Excel で開く」のワークフローが成立するため、評価エンジンなしでも実用範囲内と判断し優先度を下げる。
- グラフは既存ファイルの不明パーツ保存によりラウンドトリップのみ対応。新規作成の需要が出たら改めて検討。

---

## 2026-05-06 JST — 未対応機能の「保持だけで十分」観点の現状確認

User question: xlsx/docx/pptx の未実装機能について、ファイルを読み込んで A1 など一部だけ編集して保存した時に、Office で開いて元の内容が保持されれば十分ではないか。

### 実装確認

- `XSSFWorkbook`, `XWPFDocument`, `XMLSlideShow` はいずれも `_preservedEntries` を持つ。
- 読み込み時にモデルが理解しない ZIP entry を byte[] として退避し、保存時に preserved entries を先に書いてからモデル生成 entry で上書きする。
- 焦点テスト実行: `dotnet test tests/DotnetPoi.Core.Tests/DotnetPoi.Core.Tests.csproj --filter "FullyQualifiedName~RoundTrip_UnknownParts_Preserved"` passed。XSSF/XWPF/XSLF の 3 tests 合格。

### 判断

- 別パーツとして存在する未対応機能は、基本的に byte-for-byte で保持される可能性が高い。
  - xlsx: `xl/charts/*`, `xl/comments*.xml`, VML drawings, unsupported pivot/cache/chart-like parts。
  - docx: `word/comments.xml`, `word/footnotes.xml`, `word/endnotes.xml`, embedded OLE/media/custom parts/styles など。
  - pptx: notes slide parts, charts, media binaries, custom layouts/masters/themes, unknown animation/transition/supporting parts など。
- ただし中心 XML がモデル再生成される箇所は危険。
  - xlsx: edited/loaded worksheet XML (`xl/worksheets/sheetN.xml`) とその rels はモデルが書き直すため、未対応要素が同じ sheet XML 内に埋まっている場合は消える可能性がある。
  - docx: `word/document.xml` は再生成されるため、本文内の `w:txbxContent`, track changes (`w:ins`/`w:del`), table cell merge/borders/section columns 等はモデル対応がなければ消える可能性が高い。
  - pptx: `ppt/slides/slideN.xml` は再生成されるため、同一スライド XML 内の `p:grpSp`, `p:cxnSp`, unsupported shapes, video/audio refs 等は消える可能性が高い。
- Content Types と relationship parts もモデル生成で上書きされる場合があり、unknown part 本体が残っても参照が消えるリスクがある。特に中心 XML/中心 rels から参照される機能は実ファイル fixtures で確認が必要。

### 次にやると良い確認

実装より先に「実在 Office ファイルを読み込み、A1/本文/スライドに軽微編集して保存し、ZIP entry と rels の差分で保持確認」する fixture tests を追加するのが有効。対象は chart/comment xlsx、docx comments/footnotes/textbox/track changes/OLE、pptx notes/group/connector/media。

### POI test-data の候補ファイル

Apache POI submodule の `poi/test-data/` に実ファイルが多数ある。保持テストの初期候補:

- xlsx chart: `poi/test-data/spreadsheet/WithChart.xlsx`, `WithThreeCharts.xlsx`, `WithTwoCharts.xlsx`, `123233_charts.xlsx`
- xlsx comments: `poi/test-data/spreadsheet/SimpleWithComments.xlsx`, `comments.xlsx`
- xlsx textbox / drawing: `poi/test-data/spreadsheet/WithTextBox.xlsx`, `WithTextBox2.xlsx`
- xlsx OLE/embed: `poi/test-data/spreadsheet/WithEmbeded.xlsx`, `ExcelWithAttachments.xlsm`
- docx comments: `poi/test-data/document/comment.docx`, `testComment.docx`
- docx footnotes/endnotes: `poi/test-data/document/footnotes.docx`, `endnotes.docx`, `table_footnotes.docx`, `form_footnotes.docx`
- docx track changes: `poi/test-data/document/delins.docx`, `bug56075-changeTracking_on.docx`, `documentProtection_trackedChanges_no_password.docx`
- docx OLE/attachments: `poi/test-data/document/EmbeddedDocument.docx`, `WordWithAttachments.docx`
- docx columns: `poi/test-data/document/ThreeColHead.docx`, `ThreeColFoot.docx`, `ThreeColHeadFoot.docx`
- pptx notes/comments: `poi/test-data/slideshow/45545_Comment.pptx`, `sample_pptx_grouping_issues.pptx`
- pptx charts: `poi/test-data/slideshow/line-chart.pptx`, `bar-chart.pptx`, `pie-chart.pptx`, `scatter-chart.pptx`, `radar-chart.pptx`
- pptx media: `poi/test-data/slideshow/EmbeddedVideo.pptx`, `EmbeddedAudio.pptx`
- pptx SmartArt/group/shapes: `poi/test-data/slideshow/SmartArt.pptx`, `smartart-simple.pptx`, `sample_pptx_grouping_issues.pptx`, `shapes.pptx`, `customGeo.pptx`
- pptx OLE/attachments: `poi/test-data/slideshow/PPTWithAttachments.pptm`, `45545_Comment.pptx`

Notes:

- ZIP entry inspection confirmed representative parts exist, e.g. `xl/charts/chart1.xml`, `xl/comments1.xml`, `word/comments.xml`, `word/footnotes.xml`, `word/embeddings/*`, `ppt/media/media1.mp4`, `ppt/media/media1.mp3`, `ppt/comments/comment1.xml`, `ppt/notesSlides/*`, `ppt/charts/chart1.xml`, `ppt/diagrams/*`, `ppt/embeddings/*`。
- For preservation tests, assert both unknown part bytes and relationship/reference survival. The most fragile part is regenerated center XML/rels (`word/document.xml`, `ppt/slides/slideN.xml`, `xl/worksheets/sheetN.xml`, and their `.rels`)。

---

## 2026-05-15 JST — Phase 9 ドキュメントサイト生成 (28 pages)

- Current task: Phase 9 ドキュメントをユーザーの手を借りず自力で全ページ生成する。
- Scope: DOC_PLAN.md に沿い、HTML 生成まで完了。コミットしない。

### やったこと

**DocsGenerator 拡張:**
- Markdown パーサに **テーブル対応** 追加（`|...|` → `<table>`/`<th>`/`<td>`）
- Markdown パーサに **太字対応** 追加（`**text**` → `<strong>`）
- `site.css` にテーブルスタイル追加（borders, striped rows, accent header）

**Getting Started セクション (5 pages):**
- installation.md — NuGet インストール手順、パッケージ分割説明、動作確認
- first-workbook.md — xlsx 初めてのワークブック作成ガイド
- first-document.md — docx 初めてのドキュメント作成ガイド
- first-presentation.md — pptx 初めてのプレゼンテーション作成ガイド
- usage-samples.md — 既存

**Compatibility セクション (4 pages):**
- format-coverage.md — xlsx〜xls 全フォーマットの機能マトリクス
- limitations.md — 既知の制限事項（formula eval, charts, styles, etc.）
- package-split.md — Core vs Formula 分割設計の説明
- interop.md — Java POI 双方向互換性テストの説明

**xlsx ガイド (12 pages):**
- cell-types.md, styles.md, layout.md, images.md, formulas.md
- data-validation.md, conditional-formatting.md, auto-filter.md
- pivot-tables.md, protection.md, rich-text.md, macros.md

**docx ガイド (7 pages):**
- paragraphs.md, tables.md, images.md, headers-footers.md
- hyperlinks.md, fields.md, sections.md

**pptx ガイド (4 pages):**
- slides.md, text.md, images.md, tables.md

**xls ガイド (1 page):**
- overview.md

**Reference (1 page):**
- examples-index.md — 15の example 一覧表

### Verification
- `dotnet build tools/DotnetPoi.DocsGenerator/` — Passed
- `dotnet run --project tools/DotnetPoi.DocsGenerator -- docs_src docs` — 35 HTML files generated
- テーブル `|...|` → 正しく `<table>` HTML に変換
- 太字 `**text**` → 正しく `<strong>` HTML に変換
- ナビゲーション: 6 sections, 28 items, 35 pages total

## 2026-05-06 JST - xlsx cell alignment/wrapText status check

- User question: xlsx のセル右揃えとワードラップ機能は未実装か。
- Checked `src/DotnetPoi.Core/XSSF/UserModel/XSSFCellStyle.cs`: `setAlignment(HorizontalAlignment)` / `getAlignment()` and `setWrapText(bool)` / `getWrapText()` are implemented for XSSF.
- Checked `src/DotnetPoi.Core/XSSF/UserModel/XSSFWorkbook.cs`: style XML write emits `applyAlignment="true"` and `<alignment horizontal="right" wrapText="true">` when set; read path parses `horizontal`, `vertical`, `wrapText`, `indent`, and `textRotation`.
- Existing test coverage: `RoundTrip_StyledCell_AlignmentRestored` covers center/top/wrapText/indent/rotation, but there is no focused assertion specifically for `HorizontalAlignment.Right`.
- Note: HSSF (`.xls`) style alignment/wrap APIs are still stubs returning defaults / throwing `NotImplementedException`; this status check is for xlsx/XSSF only.

## 2026-05-06 JST - docs tagline update

- User requested adding this phrase to docs: "Apache POI-compatible un-official .NET library for Office files, focused on faithful round-trip and Java POI interop."
- Updated `docs_src/site.json` description so the phrase appears on the generated home page and as the site meta description.
- Updated `tools/DotnetPoi.DocsGenerator/Program.cs` fallback `DefaultDescription` to keep generated output consistent if `site.json` is unavailable.
- Regenerated `docs/` with `DOTNET_ROLL_FORWARD=Major dotnet run --project tools/DotnetPoi.DocsGenerator -- docs_src docs` because only .NET 9 runtime is installed locally while the docs generator targets net8.0.
- Verified with `rg` that the phrase appears in `docs/index.html`, `docs/getting-started/installation.html`, `docs_src/site.json`, and the generator fallback.

---

## 2026-05-06 JST — Round-trip + interop テスト追加（active sheet, auto filter, protection, docx fields）

### やったこと

**1. バグ修正: activeTab が常に "0" と書かれていた**

- `XSSFWorkbook.WriteWorkbook()` の `activeTab` 属性がハードコード `"0"` になっていた → `_activeSheetIndex.ToString()` に修正
- `ReadWorkbookCalcPr()` に `workbookView/activeTab` のパース追加 → round-trip が成立するようになった

**2. xlsx round-trip テスト追加 (XSSFWorkbookTests.cs)**

| テスト | 内容 |
|--------|------|
| `RoundTrip_ActiveSheetIndex_Preserved` | setActiveSheet(1) → write → XML に `activeTab="1"` → read back → getActiveSheetIndex() == 1 |
| `ActiveCellApi_WorksInMemory` | setActiveCell/getActiveCell/setSelected/isSelected の in-memory API 確認（XML 非対応） |

**3. Interop Direction B フィクスチャ生成 (WriteForPoiTests.cs) — C# → Java POI**

| テスト | 出力ファイル | 内容 |
|--------|-------------|------|
| `Write_AutoFilterWorkbook_CreatesFixtureForPoi` | `phase-autofilter.xlsx` | 3行×2列のデータにオートフィルター (A1:B3) |
| `Write_ProtectedWorkbook_CreatesFixtureForPoi` | `phase-protection.xlsx` | シート保護 + ブック保護 |
| `Write_ActiveSheetWorkbook_CreatesFixtureForPoi` | `phase-active-sheet.xlsx` | 3 シート、active=1 (Second) |
| `Write_DocxWithFields_CreatesFixtureForPoi` | `phase-docx-fields.docx` | PAGE / TOC / MERGEFIELD の 3 フィールド |

**4. Interop Direction A リーダー (ReadPoiGeneratedTests.cs) — Java POI → C#**

| テスト | 入力 | 検証内容 |
|--------|------|---------|
| `Read_AutoFilterSheet_GeneratedByPoi` | `phase-autofilter.xlsx` | AutoFilter の範囲 + セル値 |
| `Read_ProtectedSheet_GeneratedByPoi` | `phase-protection.xlsx` | シート保護 + ブック保護 + セル値 |
| `Read_ActiveSheet_GeneratedByPoi` | `phase-active-sheet.xlsx` | 3 シート、activeSheetIndex=1 |
| `Read_DocxWithFields_GeneratedByPoi` | `phase-docx-fields.docx` | PAGE / TOC / MERGEFIELD の Instruction |

**5. Java 側フィクスチャ生成 + リーダー (WriteForDotnetTest.java + ReadFromDotnetTest.java)**

- `writePhaseAutoFilterSheet()` — xlsx オートフィルター
- `writePhaseProtectedSheet()` — xlsx 保護付き
- `writePhaseActiveSheet()` — active sheet 設定
- `writePhaseDocxWithFields()` — docx フィールド
- `readPhaseAutoFilterSheet()` — C# 生成フィクスチャ検証
- `readPhaseProtectedSheet()` — C# 生成フィクスチャ検証
- `readPhaseActiveSheet()` — C# 生成フィクスチャ検証
- `readPhaseDocxFields()` — C# 生成フィクスチャ検証

### テスト結果

- **Core round-trip**: 81 passed (+2 new)
- **Interop C# WriteForPoi**: 4 passed (+4 new)
- **全ビルド**: 0 errors
- Java 側は Maven 環境で `mvn test` 実行時に対応テストが追加される

---

## 2026-05-06 JST — 全ファイルメンテナンス（preservation 確認結果反映 + カウント更新）

### やったこと

**POI test-data を使った preservation 実証テストを作成・実行:**

- `PreservationVerificationTests.cs` — 3 tests:
  - `Docx_StylesAndComments_Preserved` — 55966.docx (styles) + testComment.docx (comments) → 全 entry 保持 ✅
  - `Xlsx_RoundTrip_VerifyEntryPreservation` — 123233_charts.xlsx (4 charts + drawing) → 全 entry 保持 ✅（Sheet の大文字小文字のみ正規化）
  - `Pptx_RoundTrip_VerifyEntryPreservation` — 2411-Performance_Up.pptx (48 slides, 41 notes slides, 231 entries) → 全 entry 保持 ✅
- パス解決のバグを修正（`Path.GetDirectoryName("/workspace")` が空文字列を返す問題）
- Core.Tests に 3 tests 追加 → 231 → **234**

**各ファイルの更新:**

| ファイル | 変更内容 |
|---------|---------|
| `NOW.md` | テスト数 231→234, 296→299。docx styles 注釈を「style定義ファイルは🔵保持」に更新。日付更新 |
| `README.md` | バッジ 296→299, テスト表 231→234, 296→299 |
| `agents.md` | Core tests (228)→(234)。docx progress 行に styles/comments/footnotes/endnotes/OLE 🔵 preserve 追記 |
| `docs_src/content/compatibility/format-coverage.md` | xlsx: external data connections ❌→🔵; docx: comments/footnotes-endnotes ❌→🔵, styles 注釈更新, OLE embeddings 🔵行追加; pptx: notes slides ❌→🔵, video/audio ❌→🔵. **全6変更 + OLE追加** |
| `tools/DotnetPoi.DocsGenerator/DotnetPoi.DocsGenerator.csproj` | TargetFramework net9.0→net8.0 (SDK互換) |
| `docs/` | format-coverage.md更新を反映して再生成 (35 HTML files) |

### 確認結果（preservation verification）

- pptx 2411-Performance_Up.pptx: 231 entries **全部保持**（notes slides 41 + layouts/masters/theme/media/printerSettings）
- xlsx 123233_charts.xlsx: 20 entries **全部保持**（charts 4 + drawing）
- docx 55966.docx + testComment.docx: 全 entries **保持**（styles, comments, glossary）
- ❌ 本当に失われる: auto-shapes (drawing.xml内の`<xdr:sp>`), sparklines (sheet.xml extLst), group shapes/connectors (slide.xml内), track changes/table borders/text boxes/SDT (document.xml内)
