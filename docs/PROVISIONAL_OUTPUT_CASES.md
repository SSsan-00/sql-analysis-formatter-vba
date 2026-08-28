# 暫定実装ケース

この資料は、ユーザーレビュー前の推測期待値または現行実装出力を管理する。
2026-07-20時点で、登録済み82ケースはすべて最終レビュー済みであり、現在レビュー待ちのケースはない。

## 共通ルール

- 更新系SQLにSELECT処理が含まれる場合、SELECT表を先、最終的な`＜データ移送表＞`を後に出力する。
- SELECTがネストする場合、内側から外側へ出力する。
- 無名SELECTには`SQ1`、`SQ2`を付け、派生テーブルやCTEの既存名は維持する。
- SELECTを含まないINSERT、UPDATE、DELETEは`＜データ移送表＞`だけを出力する。

## 確定済みの最新ケース

| ケース | SQL概要 | 確定規則 |
| --- | --- | --- |
| SEL-082 | 複合条件を持つTHEN・ELSE両側のネストCASE | ANDとORを独立行に分け、先頭側の括弧グループは外側条件と同じ段落、後続グループは1段深く配置する |

SEL-082は期待値ブックと`OutputReportCases.json`へ正式登録済みである。同じ相対配置をSELECT、INSERT VALUES、INSERT SELECT、UPDATE、DELETEに適用するパーサー単体テストも登録している。

## 現在の制約

- INSERT VALUESは対象列を明示した単一行と複数行を表形式へ変換する。複数行は`＜VALUES n行目＞`ごとの移送表へ分ける。
- DEFAULT VALUES、INSERT EXECUTEは原因付きフォールバックとする。
- INSERT SELECTのUNIONとUNION ALLは各SELECTを移送パターンへ分けて出力する。EXCEPTとINTERSECTを直接ソースにする場合は原因付きフォールバックとする。
- 複合SELECTの分岐間で同じSQL別名が異なる物理テーブルを指す場合は、最初の別名を維持し、後続の束縛へ複合クエリ内で未使用の`tbN`を表示専用で割り当てる。列修飾子、条件、結合、CASE、INSERT UNIONの移送パターンに加え、SQL解析B列の別名宣言・束縛参照とC列の変換内容も同じ表示別名へ統一する。A列の元SQL、文字列、コメント、内側の同名別名スコープ、列エイリアス、物理テーブル利用情報は変更しない。

## レビュー手順

新しい暫定ケースを追加した場合は、次のコマンドで元SQLと和名定義を開発用ブックへ投入する。

```powershell
powershell -ExecutionPolicy Bypass -File tools/Set-ManualOutputCase.ps1 -CaseId <case-id>
```

`解析`を実行し、`アウトプット①`シートを確認する。暫定期待値と異なる場合は、従来どおりシートへ正しい期待値を記入する。
レビュー確定後は、`tests/SqlAnalysisFormatter.OutputExpectations.xlsx`、`ManualOutputCases.json`の`review_status`、`OutputReportCases.json`への登録、本資料を同じ変更で更新する。
