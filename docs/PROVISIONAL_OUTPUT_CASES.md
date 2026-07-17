# 暫定実装ケース

この資料は、ユーザーレビュー前の推測期待値で実装したケースを管理する。
2026-07-17時点で、次の4ケースは回帰テストへ登録済みだが、最終レビューは未実施となる。

## 共通ルール

- 更新系SQLにSELECT処理が含まれる場合、SELECT表を先、最終的な`＜データ移送表＞`を後に出力する。
- SELECTがネストする場合、内側から外側へ出力する。
- 無名SELECTには`SQ1`、`SQ2`を付け、派生テーブルやCTEの既存名は維持する。
- SELECTを含まないINSERT、UPDATE、DELETEは`＜データ移送表＞`だけを出力する。

## レビュー対象

| ケース | SQL概要 | 暫定期待値 |
| --- | --- | --- |
| SEL-068 | SUMの引数にあるCASE | 外側を`SUM(CASE結果)`と表示し、CASE分岐を右側へ複数行で出力する |
| SEL-073 | TOP内のCASE | `取得件数`を`CASE結果`として分岐を右側へ出力する |
| SEL-074 | OFFSET内のCASE | `取得範囲`のCASEを結果参照へ置換し、分岐を右側へ出力する |
| SEL-075 | 取得結果を直接返すCASE | Q列へ`CASE結果`、AF列以降へ分岐を複数行で出力し、縮小表示しない |

SQLはA5:SQL Mk-2 2.21.2の`Ctrl+Q`で実整形し、`tests/ManualOutputCases.json`へ保存している。

## 実装前の期待値レビュー

次のケースは、実装前に出力イメージを確認するための推測期待値である。`tests/SqlAnalysisFormatter.OutputExpectations.xlsx`にはレビュー用シートを追加済みだが、`tests/OutputReportCases.json`には未登録のため、76件の回帰テスト対象には含まれない。

| ケース | SQL概要 | 暫定期待値 |
| --- | --- | --- |
| SEL-076 | 複数行INSERT VALUES | `＜VALUES 1行目＞`、`＜VALUES 2行目＞`の順に、行ごとに独立したデータ移送表を出力する。各ラベルは表の外に置き、各表本体だけを外枠で囲う |

## 現在の制約

- INSERT VALUESは対象列を明示した単一行だけを表形式へ変換する。SEL-076の複数行VALUESは期待値レビュー後に実装する。
- DEFAULT VALUES、INSERT EXECUTEは原因付きフォールバックとする。
- INSERT SELECTのUNIONとUNION ALLは各SELECTを移送パターンへ分けて出力する。EXCEPTとINTERSECTを直接ソースにする場合は原因付きフォールバックとする。

## レビュー手順

次のコマンドで対象ケースの元SQLと和名定義を開発用ブックへ投入する。

```powershell
powershell -ExecutionPolicy Bypass -File tools/Set-ManualOutputCase.ps1 -CaseId SEL-068
```

`解析`を実行し、`アウトプット`シートを確認する。暫定期待値と異なる場合は、従来どおりシートへ正しい期待値を記入する。
レビュー確定後は、`tests/SqlAnalysisFormatter.OutputExpectations.xlsx`、両JSONの`review_status`、本資料を同じ変更で更新する。
