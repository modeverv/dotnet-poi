# dotnet-poi

An **unofficial**, faithful port of [Apache POI](https://poi.apache.org/) for .NET.

![License](https://img.shields.io/badge/license-Apache%202.0-blue)
![Status](https://img.shields.io/badge/status-WIP-yellow)
![Phase](https://img.shields.io/badge/phase-0%20%E2%80%94%20xlsx%20write-orange)

## Philosophy

- 🔱 Maximum fidelity to upstream Apache POI — we follow, not reinvent
- 🤖 Ported class-by-class with LLM assistance, with tests written alongside
- 💸 Free forever. No EULA. No maintenance fee. No exceptions.
- 📖 Apache POI is the source of truth — included as a git submodule
- ⚠️ Not affiliated with the Apache Software Foundation

---

## Status

### Current Phase: Phase 0 — xlsx Write (Ready to Start)

| Phase | Description | Target | Status |
|---|---|---|---|
| **-1** | **XML output parity (Java vs .NET)** | **—** | ✅ Done |
| **0** | **xlsx write (string / number)** | **v0.1** | 🚧 Ready to start |
| 1 | xlsx read | v0.2 | ⬜ Not started |
| 2 | Styles & formatting (font, color, border) | v0.3 | ⬜ Not started |
| 2.5 | Images & drawing (XSSFPicture, XSSFDrawing) | v0.35 | ⬜ Not started |
| 3 | SS common interface (IWorkbook / ISheet) | v0.4 | ⬜ Not started |
| 4 | POIFS + HSSF (xls read/write) | v0.5 | ⬜ Not started |
| 5 | Formula engine (FormulaEvaluator) | v1.0 | ⬜ Not started |
| 6 | Word / PowerPoint formats | v1.x | ⬜ Not started |

### Phase 0 — Class Progress

| Class | Ported | Tested | Notes |
|---|---|---|---|
| `XSSFWorkbook` | ⬜ | ⬜ | |
| `XSSFSheet` | ⬜ | ⬜ | |
| `XSSFRow` | ⬜ | ⬜ | |
| `XSSFCell` | ⬜ | ⬜ | Formula deferred to Phase 5 |
| `XSSFCreationHelper` | ⬜ | ⬜ | |

Legend: ✅ Done / 🚧 In Progress / ⬜ Not started

### Phase -1 Foundation

Phase -1 is complete. The project now has a `PoiXmlWriter` foundation for reproducing Apache POI/XMLBeans OOXML output at byte-level fidelity.

What is locked down:

- Java Apache POI fixture generation under `tests/DotnetPoi.Interop.Tests/fixtures/xml-parity/`
- byte-level fixture comparisons for XML declaration shape, empty element style, attribute order, namespace order, explicit zero/default attributes, element order, whitespace, and scalar formatting
- a source gate test that fails if production code bypasses `PoiXmlWriter` with direct XML APIs such as `XmlWriter`, `XDocument`, `XElement`, `XmlDocument`, or `XmlSerializer`

Before Phase 0 work is considered healthy, the XML parity tests must stay green:

```bash
dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --filter PoiXmlWriter
```

---

## Quick Start

> ⚠️ Phase 0 is still in progress. NuGet package not yet published.

```bash
git clone --recurse-submodules https://github.com/yourname/dotnet-poi
cd dotnet-poi
dotnet build
dotnet test
```

### Usage (after v0.1)

```csharp
using DotnetPoi.XSSF.UserModel;

var workbook = new XSSFWorkbook();
var sheet = workbook.CreateSheet("Sheet1");
var row = sheet.CreateRow(0);
row.CreateCell(0).SetCellValue("Hello");
row.CreateCell(1).SetCellValue(42);

using var fs = new FileStream("output.xlsx", FileMode.Create);
workbook.Write(fs);
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

### 現在のフェーズ: Phase 0 — xlsx 書き出し（開始可能）

| Phase | 内容 | バージョン目標 | 状態 |
|---|---|---|---|
| **-1** | **XML 出力挙動の統一（Java vs .NET）** | **—** | ✅ 完了 |
| **0** | **xlsx 書き出し（文字・数値）** | **v0.1** | 🚧 開始可能 |
| 1 | xlsx 読み込み | v0.2 | ⬜ 未着手 |
| 2 | スタイル・書式（フォント・色・罫線） | v0.3 | ⬜ 未着手 |
| 2.5 | 画像・図形（XSSFPicture、XSSFDrawing） | v0.35 | ⬜ 未着手 |
| 3 | SS 共通インターフェース（IWorkbook / ISheet） | v0.4 | ⬜ 未着手 |
| 4 | POIFS + HSSF（xls 読み書き） | v0.5 | ⬜ 未着手 |
| 5 | 数式エンジン（FormulaEvaluator） | v1.0 | ⬜ 未着手 |
| 6 | Word / PowerPoint 形式 | v1.x | ⬜ 未着手 |

### Phase 0 クラス別進捗

| クラス | 移植 | テスト | 備考 |
|---|---|---|---|
| `XSSFWorkbook` | ⬜ | ⬜ | |
| `XSSFSheet` | ⬜ | ⬜ | |
| `XSSFRow` | ⬜ | ⬜ | |
| `XSSFCell` | ⬜ | ⬜ | 数式は Phase 5 送り |
| `XSSFCreationHelper` | ⬜ | ⬜ | |

凡例: ✅ 完了 / 🚧 進行中 / ⬜ 未着手

### Phase -1 基盤

Phase -1 は完了しました。Apache POI/XMLBeans の OOXML 出力にバイト列レベルで寄せるための基盤として `PoiXmlWriter` を追加しています。

固定済みの内容:

- `tests/DotnetPoi.Interop.Tests/fixtures/xml-parity/` 以下の Java Apache POI 生成 fixture
- XML 宣言、空要素、属性順、namespace 順、ゼロ値・デフォルト値属性、要素順、空白、数値表現の byte-level fixture 比較
- production code が `PoiXmlWriter` を迂回して `XmlWriter`、`XDocument`、`XElement`、`XmlDocument`、`XmlSerializer` などを直接使った場合に落ちるゲートテスト

Phase 0 以降の作業では、まずこの XML parity テストが通っていることを確認します。

```bash
dotnet test tests/DotnetPoi.SS.Tests/DotnetPoi.SS.Tests.csproj --filter PoiXmlWriter
```

---

## クイックスタート

> ⚠️ 現在 Phase 0 進行中につき、まだ NuGet パッケージは公開されていません。

```bash
git clone --recurse-submodules https://github.com/yourname/dotnet-poi
cd dotnet-poi
dotnet build
dotnet test
```

### 使用例（v0.1 完成後）

```csharp
using DotnetPoi.XSSF.UserModel;

var workbook = new XSSFWorkbook();
var sheet = workbook.CreateSheet("Sheet1");
var row = sheet.CreateRow(0);
row.CreateCell(0).SetCellValue("Hello");
row.CreateCell(1).SetCellValue(42);

using var fs = new FileStream("output.xlsx", FileMode.Create);
workbook.Write(fs);
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
