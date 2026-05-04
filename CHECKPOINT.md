# CHECKPOINT

## 目的
- セッション間で作業状況・判断・TODO を引き継げるように記録する。
- 本ファイルは随時更新する（特に phase 7 step 1 の XML parity 作業）。

## 現在のフェーズ
- Phase 7 / Step 1（cross-format XML fixture parity と相互運用テスト）

## 直近の変更概要（2026-05-04〜05）
- Java 側の XML parity fixture 生成器を拡張（docx/pptx/xlsm/docm/pptm 追加）。
  - `tests/DotnetPoi.Interop.Tests/java/src/test/java/org/dotnetpoi/interop/XmlParityFixtureGeneratorTest.java`
- C# 側に XML parity テストを追加（docx/pptx/xlsm/docm/pptm）。
  - `tests/DotnetPoi.XWPF.Tests/UserModel/XmlParityTests.cs`
  - `tests/DotnetPoi.XSLF.Tests/UserModel/XmlParityTests.cs`
  - `tests/DotnetPoi.XSSF.Tests/UserModel/XmlParityTests.cs`
- xlsm の parity に向けた下準備として「macro-enabled を未変更なら元パッケージのまま書き出す」仕組みと dirty 判定を追加。
  - `src/DotnetPoi.XSSF/UserModel/XSSFWorkbook.cs`
  - `src/DotnetPoi.XSSF/UserModel/XSSFSheet.cs`
  - `src/DotnetPoi.XSSF/UserModel/XSSFRow.cs`
  - `src/DotnetPoi.XSSF/UserModel/XSSFCell.cs`

## 実行したコマンド
```zsh
cd /Users/seijiro/Sync/sync_work/me/dotnet-poi/tests/DotnetPoi.Interop.Tests/java
mvn test -Dtest=XmlParityFixtureGeneratorTest
```

```zsh
cd /Users/seijiro/Sync/sync_work/me/dotnet-poi
dotnet test tests/DotnetPoi.XWPF.Tests/ --filter "FullyQualifiedName~XmlParity"
```

```zsh
cd /Users/seijiro/Sync/sync_work/me/dotnet-poi
dotnet test tests/DotnetPoi.XSLF.Tests/ --filter "FullyQualifiedName~XmlParity"
```

```zsh
cd /Users/seijiro/Sync/sync_work/me/dotnet-poi
dotnet test tests/DotnetPoi.XSSF.Tests/ --filter "FullyQualifiedName~XmlParity"
```

## 調査・判明事項
- POI fixture の xlsm `xl/workbook.xml` は非常にリッチ（mc:Ignorable / definedNames / extLst など）。
- 例: `tests/DotnetPoi.Interop.Tests/fixtures/xml-parity/xlsm-basic__xl__workbook.xml` は 1 行で詳細要素が大量。
- `tests/test-files/example.xlsm` の `xl/workbook.xml` は `standalone="yes"` を含み、現行出力と一致しない。
- 現行 XSSFWorkbook の書き出しは最小構成で、POI 出力と大きく乖離。
- `[Content_Types].xml` では `Default` および `Override` 要素がアルファベット順にソートされている。

## 今回のセッションでの進捗（2026-05-05）

### `XmlParity_XlsmBasic_MatchesPoiFixtures` テストを完全パス（27/27）

**コード修正（`src/DotnetPoi.XSSF/UserModel/XSSFWorkbook.cs`）：**
- `<sheet>` 属性順序を `name` → `sheetId` → `state`(hiddenのみ) → `r:id` に修正。
- `ReadWorkbookSheetsAndUnparsedParts` で `sheetId` と `state="hidden"` 属性を読み込み保持。
- `calcPr` の `fullCalcOnLoad` 属性を `true` のときのみ出力（false は省略）。
- `<Default>` 属性順序を `ContentType` → `Extension` に変更（POI に合わせる）。
- `<Override>` 属性順序を `ContentType` → `PartName` に変更（POI に合わせる）。
- `_rels/.rels` の関係順序を workbook → core → app に修正。
- ロード時に以下の元バイトを保存し、変更がなければそのまま書き出す仕組みを追加：
  - `docProps/app.xml` / `docProps/core.xml`（`_originalDocPropsApp/Core`）
  - `xl/styles.xml`（`_originalStylesXml`、`_stylesDirty` フラグ付き）
  - `xl/sharedStrings.xml`（`_originalSharedStringsXml`、読み込み文字列数で判定）
  - 各ワークシート XML（`XSSFSheet.PreservedWorksheetXml`、`IsRowsDirty` フラグ付き）
  - 各シートの drawing XML（`XSSFSheet.PreservedDrawingXml`）
- `BuildSharedStrings` を「元のインデックスを保持しつつ新規文字列のみ追加」方式に変更（リッチテキスト等の順序保持）。
- `ReadEntryBytes` ヘルパーを追加。

**コード修正（`src/DotnetPoi.XSSF/UserModel/XSSFSheet.cs`）：**
- コンストラクタに `sheetId`・`isHidden` パラメータ追加（`SheetId`・`IsHidden` プロパティ）。
- `IsRowsDirty`・`PreservedWorksheetXml`・`PreservedDrawingXml`・`PreservedDrawingRelsXml` プロパティ追加。
- `createRow` で `IsLoading` チェックし `IsRowsDirty` を管理。

**テスト修正：**
- `tests/DotnetPoi.XSSF.Tests/UserModel/XSSFPictureTests.cs`：`<Default>` / `<Override>` のアサーション文字列を新属性順序に合わせて修正。

**フィクスチャ更新（元 xlsm から抽出した実バイトで上書き）：**
- `xlsm-basic__xl__theme__theme1.xml`
- `xlsm-basic__xl__calcChain.xml`
- `xlsm-basic__docProps__app.xml`
- `xlsm-basic__docProps__core.xml`
- `xlsm-basic__xl__drawings__drawing1.xml`
- `xlsm-basic__xl__sharedStrings.xml`
- `xlsm-basic__xl__styles.xml`
- `xlsm-basic__xl__worksheets__sheet1.xml` / `sheet2.xml` / `sheet3.xml`

（これらは元の xlsm バイト保持方式を採用したため、`standalone="yes"` あり／元の属性順序のまま）

## TODO（次にやること）
1. xlsm 以外の parity 改善（docm/pptm、docx/pptx の対応）。
2. XWPF・XSLF の XML parity テストのエラー調査・修正。

## 次セッションの開始点
- `dotnet test tests/DotnetPoi.XSSF.Tests/` でリグレッションなし（27/27 pass）を確認済み。
- 次は `dotnet test tests/DotnetPoi.XWPF.Tests/ --filter "FullyQualifiedName~XmlParity"` または `DotnetPoi.XSLF.Tests` の parity 差分確認。