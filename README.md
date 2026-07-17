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
- `クリア` ボタンで確認後、`変換定義` と `SQL解析` の 2 行目以降、`アウトプット` の内容をクリアし、1 行目のヘッダーを復元します。アクティブシートは変えず、3シートすべての選択セルとスクロール位置をA1へ戻します。

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
`AS` 直後の単独 ID はエイリアスとして扱い、SQL解析シートのB列では変換しません。アウトプットでは、式が参照する列のフィールドIDとエイリアスが一致する場合、そのフィールド和名を取得項目名に使用します。
`(和名未取得)` や空欄は変換対象外です。
フィールド和名には`[]`、`/`、空白、括弧、演算子、引用符、コメント記号など、T-SQLの構文として解釈される文字も使用できます。B列とアウトプットには入力した和名をそのまま表示します。
`=`、`+`、`-`、`@`で始まる和名は、Excelで数式として扱われないよう先頭へアポストロフィを付けて入力します。アポストロフィは和名には含まれません。
`#temp_users`のような一時テーブルは、A列へ`#`を含む物理テーブル名を入力すると、別名付きで参照した場合もテーブル和名をアウトプットで解決します。物理名が定義と一致しない場合は参照テーブルを`(和名未取得)`、列を元の物理名で表示します。

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
SELECT 全体は`＜DB入出力項目定義＞`、サブクエリは`サブクエリ[名前]`、更新系は`＜データ移送表＞`としてA列からCL列へ出力します。
SELECT INTOとINSERT SELECTのトップレベルSELECTは人工的なサブクエリへ分けず、`＜DB入出力項目定義＞`として出力し、各取得式を`＜データ移送表＞`の移送元へ直接対応させます。
SELECT INTO、INSERT SELECT、UPDATE、DELETEに実在するサブクエリ、CTE、派生テーブルは先に`サブクエリ[名前]`として出力し、最終表からSQ名または既存名で参照します。UPDATEの検索条件にある`EXISTS`、`IN`、比較用サブクエリも対象です。
WITH 句は名前付きサブクエリとして扱い、ネストしたサブクエリは内側から外側、最後にクエリ全体の順で出力します。
無名サブクエリには出力順に `SQ1`、`SQ2` の名前を付けます。
各表には参照テーブル、取得項目、条件、結合、並べ替えなど、文法に応じた解析結果を配置します。
SELECTの取得項目にある`tb1.*`は`tb1.全項目`、修飾のない`*`は`全項目`として表示します。SELECT INTOではデータ移送表の対象を`全項目`、移送元を`tb1.全項目`として対応させます。`TRIM(tb1.name) AS name`のように式エイリアスが参照列のフィールドIDと一致する場合は、定義済みのフィールド和名を取得項目名へ表示します。
CASEはSELECT、集計関数、WHERE、HAVING、GROUP BY、ORDER BY、JOIN、TOP、OFFSET、UPDATE SET、INSERT VALUES内から検出し、WHEN、ELSE、ネストを複数行へ展開します。複合WHEN条件は括弧とAND/ORの論理構造を再帰的に分解し、同種演算子の連続を同じ階層へ揃え、異なる論理グループを2列ずつ右へ下げます。元SQLの括弧は省略せず、各グループの先頭条件と末尾条件へ保持します。更新系のCASEはCAST、関数、演算、括弧など任意の外側式で包まれていてもCASE結果と分岐を移送方法へ表示します。外側式とTHEN/ELSEの戻り値が参照するテーブル列は出現順・重複なしの`、`区切りで移送元へ列挙し、WHEN条件だけが参照する列は除外します。
アウトプットは MS ゴシック 9 ポイント、列幅 1.14、行高 13.5、目盛り線なし、折り返しと縮小表示なしで整形します。解析結果のSQL断片にあるタブ、改行、連続空白などは半角スペース1個へ統一し、文字列リテラルと引用識別子の内部は原文どおり保持します。罫線は表本体だけを外枠で囲み、タイトル行と参照テーブル行は外枠に含めません。移送方法へ置くCASEは、分岐の行間を区切らず1つの枠で表示します。
未対応のクエリは分解せず、SQL解析シートのB列と同じ和名変換後SQLをA列へ1行ずつ出力します。1行空けた末尾に`フォールバック原因`と、原因クエリがあるアウトプットシートの行範囲を表示します。
更新系の移送表では、単純なテーブル列参照だけを`移送元`へ出力します。集計、関数、算術演算、CASTなどの式は`移送方法ほか`へ出力し、式が参照するテーブル列を出現順・重複なしの`、`区切りで`移送元`へ併記します。COUNT(*)、変数、定数など特定列を参照しない式の移送元は空欄にします。
INSERT SELECTのトップレベルSELECTにある計算式はデータ移送表の移送元または移送方法へ直接対応させ、JOIN、検索条件、グループ、集計条件は`＜DB入出力項目定義＞`へ出力します。
INSERT SELECTのトップレベルにUNIONまたはUNION ALLがある場合、SELECT表では既存の集合演算表示を使い、データ移送表では各SELECTを`＜移送パターン1＞`、`＜移送パターン2＞`のように分けて同じINSERT対象列へ対応させます。
INSERT VALUESは対象列を明示した単一行に対応します。複数行、DEFAULT VALUES、INSERT EXECUTEは原因付きでフォールバックします。

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
`run-output-golden-tests.ps1`は本体モジュールと書式回帰テスト補助を取り込み、登録済み76ケースを比較します。
保存済みの `SqlAnalysisFormatter.xlsm` にはテストモジュールを残しません。

CRUDテストケースの内容は`tests/CRUD_TEST_CASES.md`、登録済み76出力ケースは`tests/OutputReportCases.json`と`tests/SqlAnalysisFormatter.OutputExpectations.xlsx`にまとめています。ユーザーレビュー待ちの4ケースと実装前レビュー中のSEL-076は[暫定実装ケース](docs/PROVISIONAL_OUTPUT_CASES.md)で確認できます。

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
