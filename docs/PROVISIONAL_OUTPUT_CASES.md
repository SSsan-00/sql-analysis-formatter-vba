# 暫定実装ケース

この資料は、ユーザーレビュー前の推測期待値で実装したケースを管理する。
2026-07-18時点で、登録済み期待値のうち1ケースは最終レビュー未実施となる。

## 共通ルール

- 更新系SQLにSELECT処理が含まれる場合、SELECT表を先、最終的な`＜データ移送表＞`を後に出力する。
- SELECTがネストする場合、内側から外側へ出力する。
- 無名SELECTには`SQ1`、`SQ2`を付け、派生テーブルやCTEの既存名は維持する。
- SELECTを含まないINSERT、UPDATE、DELETEは`＜データ移送表＞`だけを出力する。

## レビュー対象

| ケース | SQL概要 | 暫定期待値 |
| --- | --- | --- |
| SEL-074 | OFFSET内のCASE | `取得範囲`のCASEを結果参照へ置換し、分岐を右側へ出力する |

SEL-074のSQLはA5:SQL Mk-2 2.21.2の`Ctrl+Q`で実整形している。

## 現在の制約

- INSERT VALUESは対象列を明示した単一行と複数行を表形式へ変換する。複数行は`＜VALUES n行目＞`ごとの移送表へ分ける。
- DEFAULT VALUES、INSERT EXECUTEは原因付きフォールバックとする。
- INSERT SELECTのUNIONとUNION ALLは各SELECTを移送パターンへ分けて出力する。EXCEPTとINTERSECTを直接ソースにする場合は原因付きフォールバックとする。

## レビュー手順

次のコマンドで対象ケースの元SQLと和名定義を開発用ブックへ投入する。

```powershell
powershell -ExecutionPolicy Bypass -File tools/Set-ManualOutputCase.ps1 -CaseId SEL-074
```

`解析`を実行し、`アウトプット`シートを確認する。暫定期待値と異なる場合は、従来どおりシートへ正しい期待値を記入する。
レビュー確定後は、`tests/SqlAnalysisFormatter.OutputExpectations.xlsx`、両JSONの`review_status`、本資料を同じ変更で更新する。
