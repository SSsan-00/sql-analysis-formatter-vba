# 暫定実装ケース

この資料は、ユーザーレビュー前の推測期待値で実装したケースを管理する。
2026-07-16時点で、次の6ケースは回帰テストへ登録済みだが、ハイブリッド構成の最終レビューは未実施となる。

## 共通ルール

- 更新系SQLにSELECT処理が含まれる場合、SELECT表を先、最終的な`＜データ移送表＞`を後に出力する。
- SELECTがネストする場合、内側から外側へ出力する。
- 無名SELECTには`SQ1`、`SQ2`を付け、派生テーブルやCTEの既存名は維持する。
- INSERT SELECTの移送表は、元テーブルを直接参照せず`SQn.項目名`から対象列へ移送する。
- SELECTを含まないINSERT、UPDATE、DELETEは`＜データ移送表＞`だけを出力する。

## レビュー対象

| ケース | SQL概要 | 暫定期待値 |
| --- | --- | --- |
| SEL-045 | 単純なINSERT SELECT | `サブクエリ[SQ1]`の後に3項目のデータ移送表を出力 |
| SEL-056 | JOIN、WHERE、GROUP BY、HAVINGを含むINSERT SELECT | 計算と条件を`サブクエリ[SQ1]`へ集約し、最後に5項目を移送 |
| SEL-060 | 単一行のINSERT VALUES | 変数、定数、関数を`移送方法ほか`へ記載したデータ移送表を出力 |
| SEL-062 | UPDATE SET内のTOP付きスカラーサブクエリ | SELECTを`SQ1`へ出力し、更新項目の移送元を`SQ1.金額`とする |
| SEL-063 | 派生SELECTをJOINするUPDATE | `サブクエリ[sq]`を先に出力し、JOINを含むデータ移送表を後に出力 |
| SEL-064 | EXISTS内にSELECTを持つDELETE | SELECTを`SQ1`へ出力し、削除条件を`EXISTS (SQ1)`とする |

SQLはA5:SQL Mk-2 2.21.2の`Ctrl+Q`で実整形し、`tests/ManualOutputCases.json`へ保存している。

## 現在の制約

- INSERT VALUESは対象列を明示した単一行だけを表形式へ変換する。
- 複数行VALUES、DEFAULT VALUES、INSERT EXECUTEは原因付きフォールバックとする。
- UNIONなどの集合演算を直接ソースにするINSERT SELECTは原因付きフォールバックとする。

## レビュー手順

次のコマンドで対象ケースの元SQLと和名定義を開発用ブックへ投入する。

```powershell
powershell -ExecutionPolicy Bypass -File tools/Set-ManualOutputCase.ps1 -CaseId SEL-060
```

`解析`を実行し、`アウトプット`シートを確認する。暫定期待値と異なる場合は、従来どおりシートへ正しい期待値を記入する。
レビュー確定後は、`tests/SqlAnalysisFormatter.OutputExpectations.xlsx`、両JSONの`review_status`、本資料を同じ変更で更新する。
