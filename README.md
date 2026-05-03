# dotnet-poi

An **unofficial**, faithful port of [Apache POI](https://poi.apache.org/) for .NET.

![License](https://img.shields.io/badge/license-Apache%202.0-blue)
![Status](https://img.shields.io/badge/status-WIP-yellow)
![Phase](https://img.shields.io/badge/phase-2.5%20%E2%80%94%20done%2C%20next%3A%20Phase%203-green)

## Philosophy

- 🔱 Maximum fidelity to upstream Apache POI — we follow, not reinvent
- 🤖 Ported class-by-class with LLM assistance, with tests written alongside
- 💸 Free forever. No EULA. No maintenance fee. No exceptions.
- 📖 Apache POI is the source of truth — included as a git submodule
- ⚠️ Not affiliated with the Apache Software Foundation

---

## Status

### Current Phase: Phase 3 — SS Common Interface (next up)

| Phase | Description | Target | Status |
|---|---|---|---|
| **-1** | **XML output parity (Java vs .NET)** | **—** | ✅ Done |
| **0** | **xlsx write (string / number)** | **v0.1** | ✅ Done |
| **1** | **xlsx read (string / number)** | **v0.2** | ✅ Done |
| **2** | **Styles & formatting (font, color, border)** | **v0.3** | ✅ Done |
| **2.5** | **Images & drawing (XSSFPicture, XSSFDrawing)** | **v0.35** | ✅ Done |
| 3 | SS common interface (IWorkbook / ISheet) | v0.4 | ⬜ Not started |
| 4 | POIFS + HSSF (xls read/write) | v0.5 | ⬜ Not started |
| 5 | Formula engine (FormulaEvaluator) | v1.0 | ⬜ Not started |
| 6 | Word / PowerPoint formats | v1.x | ⬜ Not started |

### Phase 0 — Class Progress

| Class | Ported | Tested | Notes |
|---|---|---|---|
| `XSSFWorkbook` | ✅ | ✅ | Minimal `.xlsx` package write |
| `XSSFSheet` | ✅ | ✅ | Minimal sheet creation and row access |
| `XSSFRow` | ✅ | ✅ | Minimal row creation and cell access |
| `XSSFCell` | ✅ | ✅ | String and numeric cells only; formulas deferred to Phase 5 |
| `XSSFCreationHelper` | ✅ | ✅ | Minimal helper instance |

Legend: ✅ Done / 🚧 In Progress / ⬜ Not started

### Phase -1 Foundation

Phase -1 is complete. The project now has a `PoiXmlWriter` foundation for reproducing Apache POI/XMLBeans OOXML output at byte-level fidelity.

What is locked down:

- Java Apache POI fixture generation under `tests/DotnetPoi.Interop.Tests/fixtures/xml-parity/`
- byte-level fixture comparisons for XML declaration shape, empty element style, attribute order, namespace order, explicit zero/default attributes, element order, whitespace, and scalar formatting
- a source gate test that fails if production code bypasses `PoiXmlWriter` with direct XML APIs such as `XmlWriter`, `XDocument`, `XElement`, `XmlDocument`, or `XmlSerializer`

For future work, the XML parity tests must stay green:

```bash
dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --filter PoiXmlWriter
```

### Phase 0 Verification

Phase 0 is complete for the first writable surface: create a workbook, create sheets/rows/cells, write string and numeric values, and save an `.xlsx` file.

Verification currently covers:

- unit tests for the Phase 0 XSSF API and generated OOXML parts
- XML parity tests for the low-level `PoiXmlWriter` fixtures captured from Apache POI/XMLBeans
- Java interop in the write direction: dotnet-poi writes an `.xlsx`, then Apache POI reads and asserts the cell values
- a runnable example under `examples/Phase0WriteExample`

Commands:

```bash
dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase0WriteExample/Phase0WriteExample.csproj
```

The example writes:

```text
examples/output/phase0-write-example.xlsx
```

Note: full `.xlsx` zip files are not expected to be byte-for-byte identical to Apache POI output because zip metadata and document timestamps can vary. Byte-level parity is asserted at the XML writer fixture layer; Phase 0 interoperability is asserted by Apache POI successfully reading dotnet-poi output.

### Phase 1 Verification

Phase 1 is complete for the minimal readable surface: open a simple `.xlsx`, restore sheets/rows/cells, and read string and numeric cell values.

Verification currently covers:

- unit tests for C# write → C# read round-trips
- Java interop in the read direction: Apache POI writes an `.xlsx`, then dotnet-poi reads and asserts the cell values
- Java interop in the write direction remains green: dotnet-poi writes an `.xlsx`, then Apache POI reads it
- a runnable interoperability example under `examples/Phase1InteropExample`

Commands:

```bash
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=WriteForDotnetTest
dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase1InteropExample/Phase1InteropExample.csproj
```

The example checks:

```text
examples/output/phase1-dotnet-poi-write.xlsx
tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase1-basic.xlsx
```

Scope note: Phase 1 covers simple `.xlsx` files with shared-string cells, numeric cells, explicit zero values, sparse cells, and multiple sheets. Formulas, styles, dates, booleans, images, and rich text remain later-phase work unless otherwise noted by tests.

### Phase 2 Verification

Phase 2 is complete for the initial practical style subset: fonts, indexed colors, foreground fills, basic borders, built-in number formats, custom number formats, and cell style references.

Verification currently covers:

- unit tests for style table generation and style readback
- Java interop in the write direction: Apache POI reads the dotnet-poi styled workbook fixture
- the XML parity writer tests remain green

Scope note: Phase 2 does not claim full style byte-level parity for every Apache POI style feature. The supported subset is semantically interoperable with Apache POI; additional style properties are tracked as Phase 2.x work.

### Phase 2.5 Verification

Phase 2.5 has its first completed slice: embedding PNG/JPEG-style image parts in `.xlsx` output through POI-compatible APIs.

Implemented surface:

- `XSSFWorkbook.addPicture(...)`
- `XSSFWorkbook.getAllPictures()`
- `XSSFCreationHelper.createClientAnchor()`
- `XSSFSheet.createDrawingPatriarch()`
- `XSSFDrawing.createPicture(...)`
- `XSSFClientAnchor`, `XSSFDrawing`, `XSSFPicture`, and `XSSFPictureData`

Verification currently covers:

- unit tests for generated media, drawing, worksheet relationship, drawing relationship, and content type parts
- C# write → C# read picture data round-trip
- Java interop in the write direction: dotnet-poi writes an image workbook, then Apache POI reads the image data and drawing shape
- a runnable example under `examples/Phase25ImagesExample`

Commands:

```bash
dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase25ImagesExample/Phase25ImagesExample.csproj
```

The example writes:

```text
examples/output/phase2_5-images-example.xlsx
```

Byte-level note: the low-level XML writer parity fixture suite is still the source of truth for POI/XMLBeans byte-level XML behavior and remains green. Full workbook ZIP byte parity is not claimed because ZIP metadata and document timestamps vary. Full byte-for-byte parity for every new drawing-related package part is also not claimed yet; current Phase 2.5 verification is semantic package interoperability with Apache POI plus continued `PoiXmlWriter` fixture parity.

---

## Quick Start

> ⚠️ NuGet package not yet published. Use a project reference or clone the repository directly.

```bash
git clone --recurse-submodules https://github.com/yourname/dotnet-poi
cd dotnet-poi
dotnet build
dotnet test
```

### Usage

```csharp
using DotnetPoi.XSSF.UserModel;

var workbook = new XSSFWorkbook();
var sheet = workbook.createSheet("Sheet1");
var row = sheet.createRow(0);
row.createCell(0).setCellValue("Hello");
row.createCell(1).setCellValue(42);

using var fs = new FileStream("output.xlsx", FileMode.Create);
workbook.write(fs);
```

Runnable example:

```bash
dotnet run --project examples/Phase0WriteExample/Phase0WriteExample.csproj
dotnet run --project examples/Phase1InteropExample/Phase1InteropExample.csproj
dotnet run --project examples/Phase25ImagesExample/Phase25ImagesExample.csproj
```

---

## Why This Project

The .NET Excel library landscape has structural problems:

- **NPOI**: Supports both xls and xlsx, but v2.8.0+ requires a commercial maintenance fee
- **ClosedXML / EPPlus**: xlsx only — cannot handle xls (BIFF format)

dotnet-poi aims to solve both problems by porting Apache POI — a battle-tested implementation — transparently and faithfully, with **no licensing strings attached, ever**.

---

## Porting Approach

Apache POI source is kept as a git submodule under `poi/`, so the original Java is always at hand. LLMs handle the mechanical Java → C# conversion; humans handle architecture decisions and quality verification.

This project is also an experiment: **can LLMs carry a large-scale, long-running intellectual porting effort?**

See [agents.md](./agents.md) for detailed porting rules.

---

## Repository Structure

```
dotnet-poi/
├── poi/                    # Apache POI submodule (read-only)
├── examples/               # Runnable examples
├── src/
│   ├── DotnetPoi.XSSF/     # xlsx (Phase 0–3)
│   ├── DotnetPoi.SS/       # Common interface (Phase 3+)
│   ├── DotnetPoi.POIFS/    # OLE2 container (Phase 4+)
│   └── DotnetPoi.HSSF/     # xls / BIFF (Phase 4+)
├── tests/
│   ├── DotnetPoi.SS.Tests/     # XML parity foundation tests
│   ├── DotnetPoi.XSSF.Tests/
│   └── DotnetPoi.Interop.Tests/ # Java/.NET fixture compatibility tests
├── tools/
│   └── porter/             # Porting progress tracker
└── agents.md               # LLM porting instructions
```

---

## Contributing

This is a personal long-term project, but PRs and Issues are welcome. Please read [agents.md](./agents.md) before contributing.

---

## License

[Apache License 2.0](./LICENSE) — same as upstream Apache POI.

---

## Disclaimer

This project is not affiliated with the Apache Software Foundation or the Apache POI project. Apache POI is a registered trademark of the Apache Software Foundation.

---
---

# dotnet-poi（日本語）

[Apache POI](https://poi.apache.org/) の **非公式** で忠実な .NET 移植です。

## 理念

- 🔱 上流の Apache POI に最大限準拠 — 独自実装ではなく追従する
- 🤖 LLM の支援によりクラス単位で移植し、テストコードも同時に作成
- 💸 永久に無料。EULA なし。メンテナンス費なし。例外なし。
- 📖 Apache POI をソースの正典として git submodule で参照
- ⚠️ Apache Software Foundation とは一切関係ありません（非公式）

---

## 対応状況

### 現在のフェーズ: Phase 3 — SS 共通インターフェース（次のフェーズ）

| Phase | 内容 | バージョン目標 | 状態 |
|---|---|---|---|
| **-1** | **XML 出力挙動の統一（Java vs .NET）** | **—** | ✅ 完了 |
| **0** | **xlsx 書き出し（文字・数値）** | **v0.1** | ✅ 完了 |
| **1** | **xlsx 読み込み（文字・数値）** | **v0.2** | ✅ 完了 |
| **2** | **スタイル・書式（フォント・色・罫線）** | **v0.3** | ✅ 完了 |
| **2.5** | **画像・図形（XSSFPicture、XSSFDrawing）** | **v0.35** | ✅ 完了 |
| 3 | SS 共通インターフェース（IWorkbook / ISheet） | v0.4 | ⬜ 未着手 |
| 4 | POIFS + HSSF（xls 読み書き） | v0.5 | ⬜ 未着手 |
| 5 | 数式エンジン（FormulaEvaluator） | v1.0 | ⬜ 未着手 |
| 6 | Word / PowerPoint 形式 | v1.x | ⬜ 未着手 |

### Phase 0 クラス別進捗

| クラス | 移植 | テスト | 備考 |
|---|---|---|---|
| `XSSFWorkbook` | ✅ | ✅ | 最小 `.xlsx` パッケージ書き出し |
| `XSSFSheet` | ✅ | ✅ | 最小のシート作成・行アクセス |
| `XSSFRow` | ✅ | ✅ | 最小の行作成・セルアクセス |
| `XSSFCell` | ✅ | ✅ | 文字列・数値セルのみ。数式は Phase 5 送り |
| `XSSFCreationHelper` | ✅ | ✅ | 最小 helper |

凡例: ✅ 完了 / 🚧 進行中 / ⬜ 未着手

### Phase -1 基盤

Phase -1 は完了しました。Apache POI/XMLBeans の OOXML 出力にバイト列レベルで寄せるための基盤として `PoiXmlWriter` を追加しています。

固定済みの内容:

- `tests/DotnetPoi.Interop.Tests/fixtures/xml-parity/` 以下の Java Apache POI 生成 fixture
- XML 宣言、空要素、属性順、namespace 順、ゼロ値・デフォルト値属性、要素順、空白、数値表現の byte-level fixture 比較
- production code が `PoiXmlWriter` を迂回して `XmlWriter`、`XDocument`、`XElement`、`XmlDocument`、`XmlSerializer` などを直接使った場合に落ちるゲートテスト

今後の作業では、この XML parity テストが通っていることを確認します。

```bash
dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --filter PoiXmlWriter
```

### Phase 0 検証

Phase 0 は、ワークブック・シート・行・セルを作成し、文字列と数値を書き込み、`.xlsx` として保存できる最初の書き出し面として完了しています。

現在の検証内容:

- Phase 0 XSSF API と生成 OOXML パーツの unit test
- Apache POI/XMLBeans から採取した fixture に対する `PoiXmlWriter` の XML byte-level parity test
- dotnet-poi が `.xlsx` を書き、Apache POI(Java) が読み取ってセル値を検証する相互運用テスト
- `examples/Phase0WriteExample` の実行サンプル

確認コマンド:

```bash
dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase0WriteExample/Phase0WriteExample.csproj
```

サンプルの出力先:

```text
examples/output/phase0-write-example.xlsx
```

注意: `.xlsx` 全体の zip ファイルは、zip metadata や document timestamp により Apache POI 出力と完全なバイト列一致にはなりません。バイト列一致は XML writer fixture 層で確認し、Phase 0 の相互運用性は Apache POI が dotnet-poi 出力を読めることで確認しています。

### Phase 1 検証

Phase 1 は、シンプルな `.xlsx` を開き、シート・行・セルを復元し、文字列セルと数値セルを読み取れる最小の読み込み面として完了しています。

現在の検証内容:

- C# 書き出し → C# 読み込みの round-trip unit test
- Apache POI(Java) が `.xlsx` を書き、dotnet-poi が読み取ってセル値を検証する相互運用テスト
- dotnet-poi が `.xlsx` を書き、Apache POI(Java) が読み取る既存の相互運用テストも green
- `examples/Phase1InteropExample` の実行サンプル

確認コマンド:

```bash
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=WriteForDotnetTest
dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase1InteropExample/Phase1InteropExample.csproj
```

サンプルが確認するファイル:

```text
examples/output/phase1-dotnet-poi-write.xlsx
tests/DotnetPoi.Interop.Tests/fixtures/from-poi/phase1-basic.xlsx
```

範囲の注意: Phase 1 が対象にしているのは shared string の文字列セル、数値セル、明示的なゼロ値、疎なセル、複数シートを含むシンプルな `.xlsx` です。数式、スタイル、日付、真偽値、画像、rich text は、テストで明示されるまでは後続フェーズの対象です。

### Phase 2 検証

Phase 2 は、実用的な最初の style subset として完了しています。対象は font、indexed color、foreground fill、basic border、built-in/custom number format、cell style reference です。

現在の検証内容:

- style table 生成と style readback の unit test
- dotnet-poi が styled workbook を書き、Apache POI(Java) が読み取る相互運用テスト
- XML parity writer test が green のまま維持されていること

注意: Apache POI の全 style feature について byte-level parity を主張する段階ではありません。現時点の保証は、実装済み subset の意味的な相互運用性です。

### Phase 2.5 検証

Phase 2.5 は、最初の完了 slice として `.xlsx` への画像埋め込みに対応しています。

実装済み API:

- `XSSFWorkbook.addPicture(...)`
- `XSSFWorkbook.getAllPictures()`
- `XSSFCreationHelper.createClientAnchor()`
- `XSSFSheet.createDrawingPatriarch()`
- `XSSFDrawing.createPicture(...)`
- `XSSFClientAnchor`、`XSSFDrawing`、`XSSFPicture`、`XSSFPictureData`

現在の検証内容:

- media / drawing / worksheet rels / drawing rels / content types の生成 unit test
- C# 書き出し → C# 読み込みで picture data を round-trip
- dotnet-poi が image workbook を書き、Apache POI(Java) が画像データと drawing shape を読み取る相互運用テスト
- `examples/Phase25ImagesExample` の実行サンプル

確認コマンド:

```bash
dotnet test tests/DotnetPoi.XSSF.Tests/DotnetPoi.XSSF.Tests.csproj --no-restore
dotnet test tests/DotnetPoi.Interop.Tests/cs/DotnetPoi.Interop.Tests.csproj --no-restore --filter Category=WriteForPoi
mvn test -f tests/DotnetPoi.Interop.Tests/java/pom.xml -Dtest=ReadFromDotnetTest
dotnet run --project examples/Phase25ImagesExample/Phase25ImagesExample.csproj
```

サンプルの出力先:

```text
examples/output/phase2_5-images-example.xlsx
```

byte-level に関する注意: Apache POI/XMLBeans の XML byte-level 挙動は、引き続き低レイヤの `PoiXmlWriter` fixture test で固定しており green です。一方で `.xlsx` 全体の zip byte-level 一致は、zip metadata や document timestamp が変わるため主張していません。また drawing 関連の全 package part について完全な byte-for-byte parity を主張する段階でもありません。Phase 2.5 の現時点の保証は、Apache POI との意味的な package 相互運用性と、`PoiXmlWriter` fixture parity の維持です。

---

## クイックスタート

> ⚠️ まだ NuGet パッケージは公開されていません。現時点では project reference か repository clone で利用してください。

```bash
git clone --recurse-submodules https://github.com/yourname/dotnet-poi
cd dotnet-poi
dotnet build
dotnet test
```

### 使用例

```csharp
using DotnetPoi.XSSF.UserModel;

var workbook = new XSSFWorkbook();
var sheet = workbook.createSheet("Sheet1");
var row = sheet.createRow(0);
row.createCell(0).setCellValue("Hello");
row.createCell(1).setCellValue(42);

using var fs = new FileStream("output.xlsx", FileMode.Create);
workbook.write(fs);
```

実行可能な example:

```bash
dotnet run --project examples/Phase0WriteExample/Phase0WriteExample.csproj
dotnet run --project examples/Phase1InteropExample/Phase1InteropExample.csproj
dotnet run --project examples/Phase25ImagesExample/Phase25ImagesExample.csproj
```

---

## なぜこのプロジェクトが必要か

.NET の Excel ライブラリには構造的な問題があります。

- **NPOI**: xls / xlsx 両対応だがv2.8.0 以降は商用利用に維持費が必要
- **ClosedXML / EPPlus**: xlsx のみ対応、xls（BIFF形式）は扱えない

dotnet-poi は Apache POI という枯れた実装を正典として透過的に移植することで、**実装品質と永続的な無償提供を両立**することを目指します。

---

## 移植方針

Apache POI のソースを `poi/` に git submodule として保持し、**常に原典を参照しながら**クラス単位で移植します。LLM が Java → C# の変換を担い、人間がアーキテクチャ判断と品質検証を行います。

これは同時に「LLM が大規模な知的作業をどこまで担えるか」という実験でもあります。

詳細な移植ルールは [agents.md](./agents.md) を参照してください。

---

## リポジトリ構造

```
dotnet-poi/
├── poi/                    # Apache POI submodule（参照専用）
├── src/
│   ├── DotnetPoi.XSSF/     # xlsx（Phase 0〜3）
│   ├── DotnetPoi.SS/       # 共通インターフェース（Phase 3〜）
│   ├── DotnetPoi.POIFS/    # OLE2コンテナ（Phase 4〜）
│   └── DotnetPoi.HSSF/     # xls / BIFF（Phase 4〜）
├── tests/
│   ├── DotnetPoi.SS.Tests/     # XML parity 基盤テスト
│   ├── DotnetPoi.XSSF.Tests/
│   └── DotnetPoi.Interop.Tests/ # Java/.NET fixture 互換テスト
├── tools/
│   └── porter/             # 移植進捗管理
└── agents.md               # LLM への移植指示
```

---

## コントリビュート

個人のライフワークプロジェクトですが、PR・Issue は歓迎します。移植に参加する場合は必ず [agents.md](./agents.md) を読んでから作業してください。

---

## ライセンス

[Apache License 2.0](./LICENSE) — 上流の Apache POI と同じです。

---

## 免責事項

このプロジェクトは Apache Software Foundation および Apache POI プロジェクトとは一切関係ありません。Apache POI は Apache Software Foundation の登録商標です。
