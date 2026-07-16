# SQL Analysis Formatter 開発ガイド

## 構成

- `src/vba/SqlAnalysisFormatter.bas`: 利用者へ配布する VBA 本体
- `src/vba/SqlAnalysisFormatterTests.bas`: 開発者用 VBA テスト
- `src/vba/SqlAnalysisFormatterGoldenTests.bas`: Excel内で書式を一括比較する回帰テスト補助
- `docs/USER_GUIDE.md`: ユーザー向けBootstrapで`README.md`として配布する利用ガイド
- `tools/SqlAnalysisFormatter.Parser`: ScriptDom を使う C# parser
- `tests/SqlAnalysisFormatter.Parser.Tests`: MSTest による C# テスト
- `tests/CRUD_TEST_CASES.md`: SQL 変換ケース資料
- `tests/OutputReportCases.json`: 登録済み75ケースの入力 SQL と和名定義
- `tests/SqlAnalysisFormatter.OutputExpectations.xlsx`: 登録済み75ケースの期待値ブック
- `tests/ManualOutputCases.json`: 確定済みケースとユーザーレビュー待ちケースの入力 SQL・和名定義
- `docs/PROVISIONAL_OUTPUT_CASES.md`: 推測期待値で実装したユーザーレビュー待ちケース
- `tools/Set-ManualOutputCase.ps1`: 指定ケースをマクロブックへ投入して期待値作成を開始するスクリプト
- `tools/run-output-golden-tests.ps1`: 実 Excel による値・書式回帰テスト

## テスト

変更後は次を実行します。

```powershell
dotnet test SqlAnalysisFormatter.sln
powershell -ExecutionPolicy Bypass -File tools/run-vba-tests.ps1
```

parser exe 経由も確認する場合は、先に publish します。

```powershell
powershell -ExecutionPolicy Bypass -File tools/publish-parser.ps1
powershell -ExecutionPolicy Bypass -File tools/run-vba-tests.ps1 -ParserExePath dist/parser/SqlAnalysisFormatter.Parser.exe
powershell -ExecutionPolicy Bypass -File tools/run-output-golden-tests.ps1
```

`run-output-golden-tests.ps1` は75ケースについてセル値、主要罫線、塗り、フォント、折り返し、縮小表示、行高、列幅、目盛り線を実 Excel で比較します。内訳はユーザーレビュー済み66件、暫定期待値9件です。
各処理の所要時間を確認する場合は`-MeasurePerformance`を付けます。

```powershell
powershell -ExecutionPolicy Bypass -File tools/run-output-golden-tests.ps1 -MeasurePerformance
```

書式は`SqlAnalysisFormatterGoldenTests.bas`によりExcel内部で比較します。最初のケースでは意図的な書式差分を検知できることも自己診断します。
機能追加は、失敗するテストを先に追加し、最小実装で成功させ、全回帰テストを維持したまま整理する TDD サイクルで進めます。

フィールド和名は、VBAがB列用SQLとは別にparser専用IDへ置換してから外部parserへ渡します。parserは描画計画の作成後に専用IDを元の和名へ一度だけ復元します。これにより、和名中のT-SQL構文文字をASTへ解釈させず、B列の表示仕様も維持します。
変換定義プロトコルはparser専用IDを含む`SAF_MAPPINGS 2`です。C#側は旧`.bas`との互換性のため`SAF_MAPPINGS 1`も読み込みます。
A列`-`のB列テーブル和名は、FROM内の実テーブルが1件かつ有効な和名が1種類の場合だけ参照テーブルのフォールバックへ使用します。複数テーブルまたは複数和名では推測しません。

利用者レビューでセル値が確定した後、共通フレームの書式だけを期待値へ反映する場合は、対象ケースを明示して次を実行します。値が一致しないケースは更新せず失敗します。

```powershell
& .\tools\run-output-golden-tests.ps1 -CaseId @('SEL-048', 'SEL-049') -RefreshFormats
```

`test-bootstrap.ps1`は、ユーザー向けREADMEに導入、初回セットアップ、トラブル対応の各セクションが含まれることも確認します。
利用手順を変更した場合は、`README.md`と`docs/USER_GUIDE.md`を同時に更新します。

## 期待値レビュー

通常のレビュー待ちケースは`ManualOutputCases.json`へSQLと和名定義だけを登録し、期待値を推測で確定しません。次のコマンドで対象ケースをブックへ投入し、利用者が`アウトプット`シートへ期待値を記入します。

```powershell
powershell -ExecutionPolicy Bypass -File tools/Set-ManualOutputCase.ps1 -CaseId SEL-048
```

レビュー後は期待値ブック、`OutputReportCases.json`、C#回帰テストへ追加し、RED・GREEN・リファクタリングの順で実装します。

利用者が推測実装を明示的に許可した場合は、失敗テストで暫定期待値を先に固定してから実装します。この場合は両JSONへ`review_status`を付け、[暫定実装ケース](PROVISIONAL_OUTPUT_CASES.md)へ概要と制約を記録します。

## Publish

parser は .NET 8.0、win-x64、self-contained、単一 exe として publish します。

```powershell
powershell -ExecutionPolicy Bypass -File tools/publish-parser.ps1
```

## Bootstrap

bootstrap 生成は利用者向けと開発者向けを分けます。

```powershell
powershell -ExecutionPolicy Bypass -File tools/build-bootstrap.ps1 -Audience User
powershell -ExecutionPolicy Bypass -File tools/build-bootstrap.ps1 -Audience Developer
powershell -ExecutionPolicy Bypass -File tools/test-bootstrap.ps1
```

生成済み bootstrap は `dist/bootstrap` に出力します。
生成物はサイズが大きいためソース管理に含めません。
