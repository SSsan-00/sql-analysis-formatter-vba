# SQL Analysis Formatter

Excel マクロで SQL の識別子を対応表の和名に変換するツールです。

導入、更新、トラブル対応の詳しい手順は[利用ガイド](docs/USER_GUIDE.md)を参照してください。

## 最短で使う

1. `SqlAnalysisFormatter.xlsm`と`SqlAnalysisFormatter.Parser.exe`を同じフォルダへ置きます。
2. Windowsにブロックされている場合は、各ファイルの`プロパティ`から`許可する`を選択します。
3. `SqlAnalysisFormatter.xlsm`を開き、配布元を確認したうえでマクロを有効化します。
4. `変換定義`シートへ定義を入力します。
5. A5:SQL Mk-2（A5M2）の`Ctrl + Q`で整形したSQLを、`SQL解析`シートのA2へ貼り付けます。
6. `解析`ボタンを押し、B列、C列以降、`アウトプット`シートを確認します。
7. 成果物を別のブックへ貼り付ける場合は、`アウトプット`シートの`コピー`ボタンを押します。

```text
SqlAnalysisFormatter/
|-- SqlAnalysisFormatter.xlsm
`-- SqlAnalysisFormatter.Parser.exe
```

parserは.NET 8を内包した単一EXEなので、通常は.NETを別途インストールする必要はありません。

## 機能

- `変換定義` シートの 2 行目以降を対応表として読み込みます。
- `SQL解析` シートの A 列 2 行目以降に入力された SQL を確認します。
- 変換後の SQL を `SQL解析` シートの B 列へ出力します。
- 変換が発生した行は、変換後の内容を `SQL解析` シートの C 列以降へ 1 件ずつ出力します。
- `アウトプット` シートへ、和名変換後クエリを AST 解析した定義表を出力します。
- `アウトプット` シートの `コピー` ボタンで、A列からCL列までの成果物をクリップボードへ登録します。
- `解析` ボタンで変換処理を実行します。
- `クリア` ボタンで確認後、`変換定義` と `SQL解析` の 2 行目以降、`アウトプット` の内容をクリアし、1 行目のヘッダーを復元します。

## シート構成

### 変換定義

| 列 | 項目 |
| --- | --- |
| A | 所属テーブルID |
| B | 所属テーブル和名 |
| C | フィールドID |
| D | フィールド和名 |

例: `tb1` / `ユーザー` / `name` / `名前`

`tb1.name` のようなテーブル修飾付き識別子は `tb1.名前` に変換されます。
所属テーブルIDは変換せず、フィールドIDだけを和名に変換します。
A 列が `-` の行は、単独のフィールド ID として和名に変換します。
参照する実テーブルが1つだけで、A列`-`のB列に有効なテーブル和名が1種類だけある場合、アウトプットの参照テーブルにはその和名だけを表示します。SQLに明示別名がある場合は`ユーザー[u]`のように別名も表示します。
`AS` 直後の単独 ID はエイリアスとして扱い、変換しません。
`(和名未取得)` や空欄は変換対象外です。
フィールド和名には`[]`、`/`、空白、括弧、演算子、引用符、コメント記号など、T-SQLの構文として解釈される文字も使用できます。B列とアウトプットには入力した和名をそのまま表示します。
`=`、`+`、`-`、`@`で始まる和名は、Excelで数式として扱われないよう先頭へアポストロフィを付けて入力します。アポストロフィは和名には含まれません。
`#temp_users`のような一時テーブルは、A列へ`#`を含む物理テーブル名を入力すると、別名付きで参照した場合もB列のテーブル和名をアウトプットで解決します。

### SQL解析

| 列 | 項目 |
| --- | --- |
| A | SQLクエリ |
| B | 和名変換後クエリ |
| C 以降 | 変換内容 |

解析対象SQLは、A5:SQL Mk-2（A5M2）の`Ctrl + Q`で整形した結果を貼り付ける前提です。
整形結果全体をコピーしてA2へ通常どおり貼り付けると、A2、A3、A4のように1行ずつ別セルへ入ります。
A列2行目以降の空白でない行を、1つのクエリとして解析します。
同じ行で複数の変換が発生した場合、C 列、D 列、E 列のように右方向へ変換後の内容だけを出力します。

### アウトプット

解析結果をもとにした別フォーマット出力用のシートです。
SELECT 全体は `＜DB入出力項目定義＞`、サブクエリは同じ表形式で `サブクエリ[名前]`、INSERT SELECT・DELETE・UPDATE は `＜データ移送表＞` として A 列から CL 列へ出力します。
SELECT INTOは複雑度にかかわらず、`サブクエリ[SQn]`、`＜DB入出力項目定義＞`、`＜データ移送表＞`の3表をこの順で出力します。
WITH 句は名前付きサブクエリとして扱い、ネストしたサブクエリは内側から外側、最後にクエリ全体の順で出力します。
無名サブクエリには出力順に `SQ1`、`SQ2` の名前を付けます。
各表には参照テーブル、取得項目、条件、結合、並べ替えなど、文法に応じた解析結果を配置します。
アウトプットは MS ゴシック 9 ポイント、列幅 1.14、行高 13.5、目盛り線なしで整形します。
未対応のクエリは分解せず、SQL解析シートのB列と同じ和名変換後SQLをA列へ1行ずつ出力します。1行空けた末尾に`フォールバック原因`と、原因クエリがあるアウトプットシートの行範囲を表示します。
更新系の移送表では、テーブル列を参照する式を`移送元`へ、変数・定数・テーブル列を参照しない関数を`移送方法ほか`へ出力します。
INSERT SELECTの計算式は、参照列を`移送元`へ、関数・計算式・列エイリアスを`移送方法ほか`へ分けて出力します。

SELECT INTOで`AS display_name`のような式エイリアスを和名表示する場合は、`変換定義`のA列を`-`、C列を`display_name`、D列を表示したい和名にします。SQL内の`AS`直後は変換せず、アウトプットの項目名だけをこの定義から解決します。

## 初回セットアップ

### マクロ付きブックを使う場合

1. `SqlAnalysisFormatter.xlsm`をダウンロードします。
2. `SqlAnalysisFormatter.Parser.exe`をブックと同じフォルダへ配置します。
3. Windowsにブロックされている場合は、両ファイルの`プロパティ`から`許可する`を選択します。
4. Excelで開き、配布元を確認したうえでマクロを有効化します。
5. `変換定義`シートに変換定義を入力します。
6. `SQL解析`シートのA2へ整形済みSQLを貼り付けます。
7. `解析`ボタンを押してB列とC列以降、`アウトプット`シートの出力を確認します。

確定フォーマットの生成には、ScriptDom を内包した `SqlAnalysisFormatter.Parser.exe` を使います。
exe が存在しない場合や実行に失敗した場合も SQL 解析は継続し、`アウトプット` には和名変換後クエリをそのまま出力します。

### ローカルBootstrapを使う場合

GitHubへ通信せずに導入するための貼り付け用 bootstrap は、開発者が生成します。
ユーザー向け bootstrap は、成果物フォルダへ `SqlAnalysisFormatter.bas`、`SqlAnalysisFormatter.Parser.exe`、利用者向け `README.md` だけを展開します。
開発者向け bootstrap は、それに加えてプロダクションコード、テストコード、開発者向けドキュメントを展開します。
展開後の成果物フォルダには bootstrap 自身のソースを含めません。

### `.bas` から導入する場合

1. 対象ブックのバックアップを作成し、`.xlsm`形式で保存します。
2. `Alt + F11`でVBAエディターを開きます。
3. `ファイル` -> `ファイルのインポート`から`src/vba/SqlAnalysisFormatter.bas`を取り込みます。
4. Excelへ戻り、`Alt + F8`から`SetupWorkbook`を実行します。
5. 3シートと、`解析`、`クリア`、`コピー`の3ボタンが作成されたことを確認し、ブックを保存します。
6. ブックと同じフォルダへ`SqlAnalysisFormatter.Parser.exe`を置きます。

`SqlAnalysisFormatter.bas`と`SqlAnalysisFormatter.Parser.exe`は連携形式が対応する同じ配布版を使用してください。

利用者が取り込む `.bas` は `src/vba/SqlAnalysisFormatter.bas` だけです。
`src/vba/SqlAnalysisFormatterTests.bas` は開発者用のテストモジュールなので、通常利用時は取り込みません。

## 困ったとき

| 症状 | 確認すること |
| --- | --- |
| シートやボタンがない | `Alt + F8`から`SetupWorkbook`を実行します。 |
| マクロを実行できない | ファイルのブロック、Excelのマクロ警告、組織のセキュリティポリシーを確認します。 |
| フィールドが変換されない | 修飾付きはA列とC列、単独フィールドはA列が`-`か確認します。 |
| アウトプットにSQLがそのまま出る | SQL末尾の`フォールバック原因`を確認します。parser未配置ならファイル名、配置場所、Windowsのブロック状態を確認します。未対応構文なら原因を添えて開発者へ連絡します。 |

詳しい確認方法とバージョン更新手順は[利用ガイド](docs/USER_GUIDE.md)にまとめています。

## 開発者向け

VBA ソースは `src/vba/SqlAnalysisFormatter.bas` で管理しています。
Excel ブックへ反映する場合は、既存の `SqlAnalysisFormatter` モジュールを削除してから再インポートしてください。
配布用ブックにはテストモジュールを含めず、本体モジュールだけを入れます。

### テスト

開発時は変更後に必ずテストを実行します。
テスト用VBAソースは `src/vba/SqlAnalysisFormatterTests.bas` です。
実Excelの書式回帰テストでは `src/vba/SqlAnalysisFormatterGoldenTests.bas` も一時ブックへ取り込みます。
Rubberduck VBA のテスト注釈を付けています。
Rubberduckを使わない場合は、次のスクリプトを実行します。

```powershell
powershell -ExecutionPolicy Bypass -File tools/run-vba-tests.ps1
```

parser exe 経由も含めて確認する場合は、先に parser を publish してから次を実行します。

```powershell
powershell -ExecutionPolicy Bypass -File tools/publish-parser.ps1
powershell -ExecutionPolicy Bypass -File tools/run-vba-tests.ps1 -ParserExePath dist/parser/SqlAnalysisFormatter.Parser.exe
powershell -ExecutionPolicy Bypass -File tools/run-output-golden-tests.ps1
```

回帰テストの処理時間を計測する場合は、`run-output-golden-tests.ps1`に`-MeasurePerformance`を付けます。

`run-vba-tests.ps1`は一時コピーしたブックに本体モジュールとVBAテストモジュールを取り込み、機能テストを実行します。
`run-output-golden-tests.ps1`は本体モジュールと書式回帰テスト補助を取り込み、確定済み60ケースを比較します。
保存済みの `SqlAnalysisFormatter.xlsm` にはテストモジュールを残しません。

CRUDテストケースの内容は `tests/CRUD_TEST_CASES.md`、確定済みの60出力ケースは `tests/OutputReportCases.json` と `tests/SqlAnalysisFormatter.OutputExpectations.xlsx` にまとめています。

### C# parser

T-SQL の AST 解析には `tools/SqlAnalysisFormatter.Parser` を使います。
対象フレームワークは `.NET 8.0` です。
問題が出た場合のみ、`.NET 9.0` 以降への更新を検討します。

```powershell
dotnet test SqlAnalysisFormatter.sln
powershell -ExecutionPolicy Bypass -File tools/publish-parser.ps1
```

publish 結果は `dist/parser/SqlAnalysisFormatter.Parser.exe` です。
単一 exe、self-contained、win-x64 として出力します。

### Bootstrap生成

貼り付け用 bootstrap は次のコマンドで生成します。

```powershell
powershell -ExecutionPolicy Bypass -File tools/build-bootstrap.ps1 -Audience User
powershell -ExecutionPolicy Bypass -File tools/build-bootstrap.ps1 -Audience Developer
powershell -ExecutionPolicy Bypass -File tools/test-bootstrap.ps1
```

生成先は `dist/bootstrap/SqlAnalysisFormatter.user.bootstrap.ps1` と `dist/bootstrap/SqlAnalysisFormatter.developer.bootstrap.ps1` です。
この生成物はサイズが大きいため、通常のソース管理対象には含めません。
