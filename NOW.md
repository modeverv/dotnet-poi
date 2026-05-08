# NOW — 現在のカバレッジ状況

## フォーマット別 機能カバレッジ一覧

凡例: ✅ 完成 / ⚠️ 一部対応（write-only など） / 🔵 不明パーツ保持で丸ごと保存はされるが作成/編集不可 / ❌ 未着手 / — 該当なし

---

### xlsx / XSSF（~78%）

| カテゴリ | 機能 | 状態 | 備考 |
|---|---|---|---|
| **セル値** | 文字列・数値・日付・論理値・エラー | ✅ | |
| **数式** | 数式テキスト書き込み + キャッシュ値読み取り | ✅ | |
| 〃 | **数式評価（計算）** | ❌ 永久凍結 | テンプレート填充 → 保存 → Excel で開くは動く。評価エンジン不在により「保存せずに計算結果をプログラムから参照」が不可 |
| **書式** | フォント（名前/サイズ/太字/斜体/色/下線/打ち消し） | ✅ | round-trip 確認済 |
| 〃 | セル塗り潰し（パターン + 前景色） | ✅ | |
| 〃 | 罫線（4辺それぞれスタイル） | ✅ | |
| 〃 | 表示形式（数値/日付/通貨 etc.） | ✅ | |
| 〃 | 配置（水平/垂直/折り返し/インデント/回転） | ✅ | |
| **レイアウト** | セル結合 / 列幅 / 行高 | ✅ | |
| 〃 | 非表示行・列 | ✅ | |
| 〃 | 固定ペイン（凍結枠） | ✅ | |
| 〃 | アクティブシート・アクティブセル | ✅ | アクティブシートは round-trip 完了。アクティブセルは in-memory API のみ |
| 〃 | 印刷設定（余白/用紙サイズ/縦横/ヘッダー・フッター） | ✅ | |
| **図形** | 画像（複数/アンカー/回転） | ✅ | |
| 〃 | ハイパーリンク | ✅ | |
| 〃 | **グラフ** | 🔵 不明パーツ保存のみ | 既存ファイルのグラフは round-trip で保持されるが新規作成不可 |
| 〃 | **コメント** | 🔵 不明パーツ保存のみ | 同上 |
| 〃 | **図形描画（オートシェイプ・図形）** | 🔵 不明パーツ保存のみ | `xdr:twoCellAnchor` の未知子要素（オートシェイプ, コネクタ, グループ）は raw XML 保持で round-trip 維持 |
| **データ** | 入力規則（データバリデーション） | ✅ | |
| 〃 | 条件付き書式 | ✅ | |
| 〃 | **フィルター / オートフィルター** | ✅ | write/read/round-trip 完了 |
| 〃 | ピボットテーブル | ⚠️ プログラム作成のみ | 新規作成は可能。既存ピボットの編集は不可だが round-trip は保持される |
| **文字列** | 共有文字列（平文） | ✅ | |
| 〃 | **リッチテキスト（文字ごとの書式）** | ✅ | XSSFRichTextString + `<rPr>` 対応完了 |
| **その他** | ブック保護 / セル保護 | ✅ | write/read/round-trip 完了 |
| 〃 | マクロ有効（xlsm） | ✅ | VBA バイト保存 + ラウンドトリップ確認済 |
| 〃 | **スパークライン** | ❌ | |
| 〃 | **外部データ接続** | 🔵 不明パーツ保存のみ | `xl/connections.xml` / `xl/externalLinks/*` は round-trip 保持されるが API モデルなし |
| 〃 | テスト数 | **166 Ooxml.Tests 全体** | XSSF/XWPF/XSLF の split test project。POI実ファイルを使った preservation / interop 系は Interop.Tests 側にも含む |

---

### docx / XWPF（~65%）

| カテゴリ | 機能 | 状態 | 備考 |
|---|---|---|---|
| **段落/ラン** | テキスト読み書き | ✅ | |
| 〃 | フォント（太字/斜体/下線/打ち消し/名前/サイズ/色） | ✅ | round-trip 確認済 |
| 〃 | 配置（左/中央/右/両端） | ✅ | |
| 〃 | インデント（左/右/一行目/ぶら下げ） | ✅ | |
| 〃 | 行間・段落前後スペース | ✅ | |
| 〃 | 箇条書き（箇条/番号付き） | ✅ | OOXML numbering 対応 |
| **表** | 表作成・読み取り（セル・行・列） | ✅ | round-trip 確認済 |
| 〃 | セル結合・罫線 | 🔵 | `tcPr`/`tblPr` の未知子要素として raw XML 保持。API からの新規作成は未対応 |
| **セクション** | ページ設定（サイズ/余白/縦横） | ✅ | |
| 〃 | ヘッダー・フッター | ✅ | round-trip 確認済。画像・書式などリッチコンテンツも `_preservedEntries` で保持（API非経由時） |
| 〃 | **段組** | ✅ | `setColumns()`/`getColumnCount()`/`getColumnSpacing()` API、round-trip 確認済 |
| **リンク** | ハイパーリンク（外部URL） | ✅ | round-trip 確認済 |
| **画像** | 画像埋め込み（インライン・回転） | ✅ | |
| 〃 | **フローティング画像（アンカー付き）** | 🔵 | `<wp:anchor>` は raw XML capture/re-emission で round-trip 保持 |
| 〃 | **テキストボックス（w:txbxContent）** | ❌ | Word の「テキストボックス」内のテキストは読めない |
| **注釈** | **コメント** | 🔵 不明パーツ保存のみ | 既存コメントは round-trip 保持されるが API モデルなし |
| 〃 | **脚注・文末脚注** | 🔵 不明パーツ保存のみ | `word/footnotes.xml` / `word/endnotes.xml` は round-trip 保持されるが API モデルなし |
| **フィールド** | **TOC（目次）/ ページ番号 / 差し込み印刷** | ✅ | write/read/round-trip 完了 |
| **SDT** | **コンテンツコントロール** | 🔵 | ブロックレベル `w:sdt`（bodyの直下）もインライン `w:sdt`（段落内）も raw XML 補完で round-trip 維持。テキストボックス（`w:txbxContent`）は DrawingML の深いネスト内にあり未対応 |
| **スタイル** | **段落スタイル（pStyle参照）** | ✅ | `setStyle()`/`getStyleID()` API、round-trip 確認済。文字/テーブルスタイルは ❌。`word/styles.xml` は🔵保持＋新規文書にデフォルトスタイル自動生成 |
| **変更履歴** | **トラックチェンジ** | ❌ | |
| **その他** | マクロ有効（docm） | ✅ | VBA バイト保存 |
| 〃 | 未知パーツ保存 | ✅ | _preservedEntries 機構実装済 |
| 〃 | **OLE 埋め込み** | 🔵 不明パーツ保存のみ | `word/embeddings/*` は round-trip 保持される |
| 〃 | テスト数 | **151 Ooxml.Tests 全体** | XSSF/XWPF/XSLF の split test project 内で検証 |

---

### pptx / XSLF（~40%）

| カテゴリ | 機能 | 状態 | 備考 |
|---|---|---|---|
| **スライド** | スライド作成・読み取り | ✅ | |
| 〃 | スライドサイズ | ✅ | |
| 〃 | **ノートスライド** | 🔵 不明パーツ保存のみ | `ppt/notesSlides/notesSlide*.xml` は round-trip 保持されるが作成不可 |
| **テキスト** | テキストボックス（p:sp）作成・読み取り | ✅ | round-trip 確認済 |
| 〃 | 複数段落 | ✅ | |
| 〃 | ラン書式（太字/斜体/下線/打ち消し/サイズ/フォント名/色） | ✅ | |
| **図形** | 画像（場所・サイズ・回転） | ✅ | round-trip 確認済 |
| 〃 | 表（p:graphicFrame / a:tbl） | ✅ | round-trip 確認済 |
| 〃 | **グループ化** | 🔵 spTree 内の未知要素を raw XML 保持 | |
| 〃 | **コネクタ・線** | 🔵 spTree 内の未知要素を raw XML 保持 | |
| 〃 | **SmartArt** | 🔵 不明パーツ保存のみ | |
| 〃 | **グラフ** | 🔵 不明パーツ保存のみ | |
| 〃 | **オートシェイプ（矩形/円/etc.）以外の図形** | ❌ | |
| **メディア** | **動画・音声埋め込み** | 🔵 不明パーツ保存のみ | 画像以外の `ppt/media/*` は round-trip 保持されるが API モデルなし |
| **アニメ** | **アニメーション / トランジション** | 🔵 不明パーツ保存のみ | |
| **テーマ** | レイアウト・マスター保存 | ✅ | テンプレート round-trip で master/layout 全保持 |
| 〃 | **レイアウト読み取り・選択** | ✅ | `getSlideLayouts()` で名前/type 取得。`createSlide(layout)` でレイアウト指定。テンプレート pptx の軽編集に対応 |
| **その他** | 未知パーツ保存 | ✅ | |
| 〃 | マクロ有効（pptm） | ✅ | VBA バイト保存 |
| 〃 | テスト数 | **151 Ooxml.Tests 全体** | XSSF/XWPF/XSLF の split test project 内で検証 |

---

### xls / HSSF（~35%）

| カテゴリ | 機能 | 状態 | 備考 |
|---|---|---|---|
| **セル値** | 文字列・数値・論理値・空白・エラー | ✅ | BIFF8 の LabelSST/Number/BoolErr/Blank 読み書きと round-trip 確認済 |
| **シート** | 複数シート / 疎な行・セル / 高列インデックス | ✅ | |
| **書式** | フォント（名前/サイズ/太字/斜体/色） | ⚠️ 一部対応 | HSSFFont/HSSFCellStyle の基本 round-trip 対応。色は palette index ベース |
| 〃 | セルスタイル（表示形式/配置/折り返し/罫線/塗り潰し） | ⚠️ 一部対応 | 主要 XF / FormatRecord の読み書き対応。完全な BIFF style parity ではない |
| **レイアウト** | 列幅 / 行高 / 非表示行・列 | ✅ | |
| 〃 | セル結合 | ✅ | MergeCells record 対応 |
| 〃 | 固定ペイン（凍結枠） | ✅ | Pane record 対応 |
| **数式** | 数式テキスト + キャッシュ値読み取り | ⚠️ 読み取り中心 | 既存 POI fixture の formula/cached result 読み取りあり。新規 BIFF formula token 書き込みは未対応 |
| 〃 | **数式評価（計算）** | ❌ | HSSFFormulaEvaluator は未移植 |
| **互換性** | POI 代表 `.xls` fixture 読み込み | ✅ | empty/Simple/SampleSS/DateFormats/styles/formula/hyperlink/comments/drawings/images/macro 等 20 fixture のロード確認 |
| 〃 | Java POI 双方向 interop | ⚠️ 一部対応 | basic/styles/layout/unicode/comprehensive fixture を C#↔Java で検証 |
| **Preservation** | OLE2 非 Workbook stream 保存 | ✅ | non-workbook stream と directory metadata を保持 |
| 〃 | マクロ付き `.xls` の VBA stream 保存 | ✅ | `_VBA_PROJECT_CUR/VBA/*` を保持 |
| 〃 | 未知 BIFF record 保存 | ✅ | 軽編集時も global/sheet unknown record を保持 |
| **未対応** | 画像/図形/グラフ/コメント/ハイパーリンク/API編集/フィルター/ピボット | ❌ | 既存 fixture のロード対象には含むが、ユーザーモデルとしての作成・編集は未実装 |
| **その他** | テスト数 | **221 Legacy.Tests 全体 + interop** | HSSF/HWPF/HSLF の split test project。POI fixture も代表ケースを含む |

---

### doc / HWPF（~20%）

| カテゴリ | 機能 | 状態 | 備考 |
|---|---|---|---|
| **読み込み** | OLE2 `.doc` オープン / FIB 解析 | ✅ | WordDocument + 0Table/1Table 選択、fallback、CLX 範囲検証あり |
| **本文** | main document text 抽出 | ✅ | CLX / piece table から本文文字列を復元 |
| 〃 | 圧縮/Unicode text piece | ✅ | `FcCompressed` 相当の compressed/uncompressed piece 読み取り |
| **UserModel** | Range / Paragraph / CharacterRun | ⚠️ 一部対応 | paragraph/run 分割、offset、text composition を確認済 |
| 〃 | 文字書式読み取り | ⚠️ 一部対応 | CHPX/CHPFKP 由来の font name/size/bold/italic/underline/strike の一部 |
| 〃 | 段落プロパティ読み取り | ⚠️ 最小対応 | justification など一部 PAPX のみ |
| 〃 | StyleSheet / FontTable | ⚠️ 最小対応 | Normal style fallback 用の最低限 |
| **編集** | no-op write / round-trip | ✅ | 代表 fixture の stream/storage を byte-for-byte 保持 |
| 〃 | 本文追記 / 文字列置換 | ⚠️ 限定対応 | main body を単一 Unicode piece として再構築。複雑な既存構造の編集エンジンではない |
| **Preservation** | OLE stream/storage / 埋め込み OLE | ✅ | ObjectPool、Data stream、未編集 stream/storage を保持 |
| **互換性** | Java POI interop | ⚠️ 一部対応 | dotnet-poi no-op saved `.doc` を Java POI が読む Direction B fixture あり |
| **未対応** | 表/画像/API編集/ヘッダー・フッター/脚注/コメント/フィールドのモデル化 | ❌ | 既存 stream は保持対象だが、HWPF usermodel としての作成・編集は未移植 |
| **その他** | テスト数 | **221 Legacy.Tests 全体 + interop** | POI `document/` fixtures 複数を対象 |

---

### ppt / HSLF（~12%）

| カテゴリ | 機能 | 状態 | 備考 |
|---|---|---|---|
| **読み込み** | OLE2 `.ppt` オープン / stream inventory | ✅ | PowerPoint Document, Current User, summary streams 検出 |
| **レコードツリー** | record header (recType/recInstance/recLen/offset/raw bytes) | ✅ | container/atom 階層を保持。unknown record を保持 |
| **スライド** | slide count / slide order (persist pointer) | ✅ | 13 fixture survey 完了。incorrect_slide_order.ppt で正しい順序 |
| **テキスト** | TextCharsAtom / TextBytesAtom 抽出 | ✅ | UTF-16LE / CP1252 対応。title/body 分離 (TextHeaderAtom) |
| 〃 | 中国語・空テキストボックス・複数テキストブロック | ⚠️ 一部対応 | 例外は出さない。文字化けや欠損は fixture 依存で要改善 |
| **編集** | no-op write / round-trip | ✅ | CompoundFile.Write で OLE2 を byte-for-byte 保持 |
| **Preservation** | OLE stream/storage / 画像 / コメント / OLE | ✅ | RoundTrip_NoOpWrite_PreservesSpecialStreams で確認 |
| **互換性** | C# round-trip tests | ✅ | 124 HSLF test methods (survey + stream + record + persist + text + write) |
| 〃 | Java POI interop Direction B | ⚠️ C#側完了 | `Write_Phase15HslfNoOp_CreatesFixtureForPoi` 追加済。Java 側 test は未追加 |
| **未対応** | スライド作成/図形編集/画像追加/ヘッダーフッター/アニメーション | ❌ | 編集は no-op 保存のみ。新規作成は不可 |

---

## 実用上「痛い」欠損分析（優先度順）

### 🔴 最優先 — 実用上致命的

| # | 欠損 | フォーマット | なぜ痛いか |
|---|---|---|---|
| 1 | **数式の評価** | xlsx | テンプレート填充 → 保存 → Excel で開く のワークフローは **問題なく動く**（数式テキストは保持され、Excel が開き直し時に再計算する）。ただし「保存せずにプログラムから計算結果を参照したい」場合や「REST API の応答として計算済み xlsx を返したい」場合は評価エンジンが必要になる。`setCellFormula` だけでキャッシュ値を設定せずに保存した場合、`<v>` 要素は省略されるが、これは valid な OOXML であり Excel は自動再計算する。「修復しますか？」は出ない。 |
| 2 | **グラフの作成** | xlsx, pptx | ほとんどすべての報告書・資料にグラフが含まれる。既存ファイルのラウンドトリップのみ可能で新規作成が一切できない。pptx は「グラフ入りスライドを生成」が主要ユースケースの一つ。 |
| 3 | **セル保護 / ブック保護** | xlsx | ✅ 完了 — sheet.xml に `<sheetProtection>`、workbook.xml に `<workbookProtection>` を書き出し、読み戻し可能。パスワードハッシュは未実装だが保護の有無はラウンドトリップする。 |
| 4 | **フィールド（TOC/ページ番号/差し込み印刷）** | docx | ✅ 完了 — XWPFField クラス、fldChar（begin/separate/end）の読み書き対応済み。addField/GetFields API あり。 |

### 🟡 重度 — 使う人次第で苦しい

| # | 欠損 | フォーマット | なぜ痛いか |
|---|---|---|---|
| 5 | **オートフィルター** | xlsx | ✅ 完了 — sheet.xml に `<autoFilter ref="...">` 要素を書き出し、読み戻し可能。フィルター条件は簡略化（ON/OFF のみ）。 |
| 6 | **スタイル（段落スタイル・テーマ）** | docx | 文書の一貫した書式設定はスタイルが要。Word 標準スタイル（Normal/Heading1/Title）を使った文書を読むと書式が失われる。 |
| 7 | **コメント** | xlsx, docx | レビューワークフローで多用される。既存コメントは不明パーツ保存で維持される可能性があるが新規作成不可。 |
| 8 | **テキストボックス（w:txbxContent）** | docx | Word のテキストボックス内のテキストが読めない。レイアウト崩れの原因に。 |
| 9 | **セル結合・罫線** | docx 表 | docx の表でセル結合ができないのは実用範囲を狭める。 |

### 🔵 軽度 — 当面許容範囲

| # | 欠損 | 理由 |
|---|---|---|
| 10 | SmartArt / アニメーション / トランジション | pptx で新規作成できないが、不明パーツ保存で維持される。プログラマティックに生成するユースケースは稀。 |
| 11 | ppt（HSLF）と xls/doc の未対応深部 | HSLF は open/text extraction/no-op preservation/round-trip まで対応済。スライド作成・図形編集・画像追加は未対応。HSSF/HWPF も基礎対応は進んだが、画像・図形・高度な書式・完全編集はまだ重い。新規開発はほぼ OOXML で行われる。 |
| 12 | 変更履歴 / スパークライン | 特殊用途または複雑すぎる。最初の round-trip 対象としては重い。 |

---

## テスト数

| プロジェクト | テスト数 | 備考 |
|---|---|---|
| Common.Tests | 79 | shared SS / XML writer / utilities |
| POIFS.Tests | 11 | OLE2 container |
| Ooxml.Tests | 166 | XSSF / XWPF / XSLF |
| Legacy.Tests | 221 | HSSF / HWPF / HSLF |
| Formula.Tests | 11 | 限定 formula evaluator subset |
| All.Tests | 7 | 全体 smoke test |
| Interop.Tests (C#) | 72 | 双方向 interop fixture 検証 (70 pass + 2 skip) |
| **Total (C#)** | **567** | 565 pass + 2 skip |
| Java POI 側 (Maven) | 44 tests | うち dotnet-poi 関連 24 tests |

---

*最終更新: 2026-05-08 (phase 17-7: XSLF layout/master 最小操作 追加)*
