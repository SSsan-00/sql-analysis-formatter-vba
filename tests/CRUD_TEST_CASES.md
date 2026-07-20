# CRUD Test Cases

機能拡張時のラバーダック用テストケースです。
変換定義とSQLを固定し、解析結果と将来の `アウトプット` シート出力を比較しながら実装します。

解析対象SQLは、A5:SQL Mk-2（A5M2）の `Ctrl+q` で整形した結果を `SQL解析` シートのA列セルへ貼り付ける前提です。
テストコードではA5M2が付ける行末スペースを `TS(...)` で明示して保持します。この文書では視認性のため行末スペースを省略しています。

## 変換定義

| 所属テーブルID | 所属テーブル和名 | フィールドID | フィールド和名 |
| --- | --- | --- | --- |
| users | ユーザー | user_id | ユーザーID |
| users | ユーザー | name | 氏名 |
| users | ユーザー | email | メール |
| orders | 注文 | order_id | 注文ID |
| orders | 注文 | user_id | 注文ユーザーID |
| orders | 注文 | amount | 金額 |
| order_items | 注文明細 | order_id | 明細注文ID |
| order_items | 注文明細 | product_id | 商品ID |
| order_items | 注文明細 | quantity | 数量 |
| users | ユーザー | manager_id | 管理者ID |
| manager | 管理者 | user_id | ユーザーID |
| manager | 管理者 | name | 氏名 |
| manager | 管理者 | status | 状態 |
| - |  | status | 状態 |
| - |  | created_at | 作成日時 |
| - |  | updated_at | 更新日時 |

## SELECT

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    trim(users.name) as name
    , users.user_id
    , orders.order_id
    , status
from
    users
    inner join orders
        on users.user_id = orders.user_id
where
    status = 'ACTIVE'
```

期待する和名変換後

```sql
select
    trim(users.氏名) as name
    , users.ユーザーID
    , orders.注文ID
    , 状態
from
    users
    inner join orders
        on users.ユーザーID = orders.注文ユーザーID
where
    状態 = 'ACTIVE'
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.氏名 |
| 2 | users.ユーザーID |
| 3 | orders.注文ID |
| 4 | 状態 |
| 5 | orders.注文ユーザーID |

## INSERT

入力（A5M2 Ctrl+Q 整形後）

```sql
insert
into orders(order_id, user_id, amount, status, created_at)
select
    orders.order_id
    , users.user_id
    , orders.amount
    , status
    , created_at
from
    users
```

期待する和名変換後

```sql
insert
into orders(order_id, user_id, amount, 状態, 作成日時)
select
    orders.注文ID
    , users.ユーザーID
    , orders.金額
    , 状態
    , 作成日時
from
    users
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | 状態 |
| 2 | 作成日時 |
| 3 | orders.注文ID |
| 4 | users.ユーザーID |
| 5 | orders.金額 |

## UPDATE

入力（A5M2 Ctrl+Q 整形後）

```sql
update users
set
    users.name = 'Taro'
    , updated_at = CURRENT_TIMESTAMP
    , status = 'ACTIVE'
where
    users.user_id = :user_id
```

期待する和名変換後

```sql
update users
set
    users.氏名 = 'Taro'
    , 更新日時 = CURRENT_TIMESTAMP
    , 状態 = 'ACTIVE'
where
    users.ユーザーID = :user_id
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.氏名 |
| 2 | 更新日時 |
| 3 | 状態 |
| 4 | users.ユーザーID |

## DELETE

入力（A5M2 Ctrl+Q 整形後）

```sql
delete
from
    orders
where
    orders.order_id in (
        select
            order_items.order_id
        from
            order_items
        where
            order_items.product_id = :product_id
    )
    and status = 'CANCELLED'
```

期待する和名変換後

```sql
delete
from
    orders
where
    orders.注文ID in (
        select
            order_items.明細注文ID
        from
            order_items
        where
            order_items.商品ID = :product_id
    )
    and 状態 = 'CANCELLED'
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.注文ID |
| 2 | order_items.明細注文ID |
| 3 | order_items.商品ID |
| 4 | 状態 |

## 複合SELECT CASE / WHEN / ELSE / HAVING / ORDER BY

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , case
        when sum(orders.amount) > 100000
            then 'VIP'
        when sum(orders.amount) between 50000 and 100000
            then 'STANDARD'
        else status
        end as rank_name
from
    users
    left join orders
        on users.user_id = orders.user_id
where
    (
        (status = 'ACTIVE' and orders.amount > 0)
        or (
            status = 'PENDING'
            and exists (
                select
                    1
                from
                    order_items
                where
                    order_items.order_id = orders.order_id
                    and order_items.quantity > 1
            )
        )
    )
group by
    users.user_id
    , status
having
    count(orders.order_id) > 0
order by
    users.user_id
    , status
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , case
        when sum(orders.金額) > 100000
            then 'VIP'
        when sum(orders.金額) between 50000 and 100000
            then 'STANDARD'
        else 状態
        end as rank_name
from
    users
    left join orders
        on users.ユーザーID = orders.注文ユーザーID
where
    (
        (状態 = 'ACTIVE' and orders.金額 > 0)
        or (
            状態 = 'PENDING'
            and exists (
                select
                    1
                from
                    order_items
                where
                    order_items.明細注文ID = orders.注文ID
                    and order_items.数量 > 1
            )
        )
    )
group by
    users.ユーザーID
    , 状態
having
    count(orders.注文ID) > 0
order by
    users.ユーザーID
    , 状態
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | orders.金額 |
| 3 | 状態 |
| 4 | orders.注文ユーザーID |
| 5 | order_items.明細注文ID |
| 6 | orders.注文ID |
| 7 | order_items.数量 |

## 自己結合

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , users.name
    , manager.name as manager_name
from
    users
    inner join users manager
        on users.manager_id = manager.user_id
where
    manager.status = status
order by
    manager.name
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , users.氏名
    , manager.氏名 as manager_name
from
    users
    inner join users manager
        on users.管理者ID = manager.ユーザーID
where
    manager.状態 = 状態
order by
    manager.氏名
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | users.氏名 |
| 3 | manager.氏名 |
| 4 | users.管理者ID |
| 5 | manager.ユーザーID |
| 6 | manager.状態 |
| 7 | 状態 |

## SELECT-INTO

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , users.email
    , status
into user_export
from
    users
where
    users.email is not null
    and status in ('ACTIVE', 'LOCKED')
order by
    users.email
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , users.メール
    , 状態
into user_export
from
    users
where
    users.メール is not null
    and 状態 in ('ACTIVE', 'LOCKED')
order by
    users.メール
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | users.メール |
| 3 | 状態 |

## UPDATE-FROM

入力（A5M2 Ctrl+Q 整形後）

```sql
update orders
set
    orders.amount = orders.amount * 1.1
    , updated_at = CURRENT_TIMESTAMP
from
    orders
    inner join users
        on orders.user_id = users.user_id
where
    (users.email like :domain or status = 'PENDING')
    and (
        orders.amount > 1000
        or exists (
            select
                1
            from
                order_items
            where
                order_items.order_id = orders.order_id
        )
    )
```

期待する和名変換後

```sql
update orders
set
    orders.金額 = orders.金額 * 1.1
    , 更新日時 = CURRENT_TIMESTAMP
from
    orders
    inner join users
        on orders.注文ユーザーID = users.ユーザーID
where
    (users.メール like :domain or 状態 = 'PENDING')
    and (
        orders.金額 > 1000
        or exists (
            select
                1
            from
                order_items
            where
                order_items.明細注文ID = orders.注文ID
        )
    )
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.金額 |
| 2 | 更新日時 |
| 3 | orders.注文ユーザーID |
| 4 | users.ユーザーID |
| 5 | users.メール |
| 6 | 状態 |
| 7 | order_items.明細注文ID |
| 8 | orders.注文ID |

## DELETE EXISTS

入力（A5M2 Ctrl+Q 整形後）

```sql
delete
from
    order_items
where
    exists (
        select
            1
        from
            orders
        where
            orders.order_id = order_items.order_id
            and (status = 'CANCELLED' or orders.amount <= 0)
    )
```

期待する和名変換後

```sql
delete
from
    order_items
where
    exists (
        select
            1
        from
            orders
        where
            orders.注文ID = order_items.明細注文ID
            and (状態 = 'CANCELLED' or orders.金額 <= 0)
    )
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.注文ID |
| 2 | order_items.明細注文ID |
| 3 | 状態 |
| 4 | orders.金額 |

## T-SQL独立ケース

T-SQLで使用する関数・構文ごとに、A5M2 `Ctrl+q` 整形後の入力と和名変換後を独立した行として検証します。
`FORMAT` は複数のフォーマット指定子、`CAST` は複数の型を同一ケース内にまとめています。

### TRIM FROM

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    trim('.' from users.name) as trimmed_name
from
    users
where
    users.user_id = @user_id
```

期待する和名変換後

```sql
select
    trim('.' from users.氏名) as trimmed_name
from
    users
where
    users.ユーザーID = @user_id
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.氏名 |
| 2 | users.ユーザーID |

### IN

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , status
from
    users
where
    status in ('ACTIVE', 'LOCKED', 'PENDING')
    and users.user_id in (
        select
            orders.user_id
        from
            orders
        where
            orders.amount > 0
    )
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , 状態
from
    users
where
    状態 in ('ACTIVE', 'LOCKED', 'PENDING')
    and users.ユーザーID in (
        select
            orders.注文ユーザーID
        from
            orders
        where
            orders.金額 > 0
    )
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | 状態 |
| 3 | orders.注文ユーザーID |
| 4 | orders.金額 |

### COALESCE

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , coalesce(users.email, users.name, 'unknown') as contact_text
from
    users
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , coalesce(users.メール, users.氏名, 'unknown') as contact_text
from
    users
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | users.メール |
| 3 | users.氏名 |

### FORMAT

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    orders.order_id
    , format(orders.amount, 'N2', 'ja-JP') as amount_n2
    , format(created_at, 'yyyy/MM/dd') as created_date
    , format(created_at, 'yyyyMMddHHmmss') as created_stamp
    , format(orders.amount, 'C', 'ja-JP') as amount_currency
from
    orders
```

期待する和名変換後

```sql
select
    orders.注文ID
    , format(orders.金額, 'N2', 'ja-JP') as amount_n2
    , format(作成日時, 'yyyy/MM/dd') as created_date
    , format(作成日時, 'yyyyMMddHHmmss') as created_stamp
    , format(orders.金額, 'C', 'ja-JP') as amount_currency
from
    orders
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.注文ID |
| 2 | orders.金額 |
| 3 | 作成日時 |

### WITH

入力（A5M2 Ctrl+Q 整形後）

```sql
with target_users as (
    select
        users.user_id
    from
        users
    where
        status = 'ACTIVE'
)
select
    target_users.user_id
from
    target_users
```

期待する和名変換後

```sql
with target_users as (
    select
        users.ユーザーID
    from
        users
    where
        状態 = 'ACTIVE'
)
select
    target_users.user_id
from
    target_users
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | 状態 |

### CAST

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    cast(users.user_id as int) as user_id_int
    , cast(orders.amount as decimal (18, 2)) as amount_decimal
    , cast(created_at as date) as created_date
    , cast(updated_at as datetime2(3)) as updated_at_dt
    , cast(status as nvarchar(20)) as status_text
from
    users
    inner join orders
        on users.user_id = orders.user_id
```

期待する和名変換後

```sql
select
    cast(users.ユーザーID as int) as user_id_int
    , cast(orders.金額 as decimal (18, 2)) as amount_decimal
    , cast(作成日時 as date) as created_date
    , cast(更新日時 as datetime2(3)) as updated_at_dt
    , cast(状態 as nvarchar(20)) as status_text
from
    users
    inner join orders
        on users.ユーザーID = orders.注文ユーザーID
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | orders.金額 |
| 3 | 作成日時 |
| 4 | 更新日時 |
| 5 | 状態 |
| 6 | orders.注文ユーザーID |

### ISNULL

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , isnull(users.email, 'unknown') as email_text
    , isnull(status, 'UNKNOWN') as status_text
from
    users
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , isnull(users.メール, 'unknown') as email_text
    , isnull(状態, 'UNKNOWN') as status_text
from
    users
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | users.メール |
| 3 | 状態 |

### SUBSTRING

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , substring(users.email, 1, 3) as email_prefix
from
    users
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , substring(users.メール, 1, 3) as email_prefix
from
    users
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | users.メール |

### ROUND

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    orders.order_id
    , round(orders.amount, 0) as amount_round0
    , round(orders.amount, 2, 1) as amount_truncate2
from
    orders
```

期待する和名変換後

```sql
select
    orders.注文ID
    , round(orders.金額, 0) as amount_round0
    , round(orders.金額, 2, 1) as amount_truncate2
from
    orders
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.注文ID |
| 2 | orders.金額 |

### SUM

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    orders.user_id
    , sum(orders.amount) as total_amount
from
    orders
group by
    orders.user_id
having
    sum(orders.amount) > 0
```

期待する和名変換後

```sql
select
    orders.注文ユーザーID
    , sum(orders.金額) as total_amount
from
    orders
group by
    orders.注文ユーザーID
having
    sum(orders.金額) > 0
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.注文ユーザーID |
| 2 | orders.金額 |

### REPLACE

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , replace (users.email, '@old.example', '@new.example') as normalized_email
from
    users
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , replace (users.メール, '@old.example', '@new.example') as normalized_email
from
    users
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | users.メール |

### DATEADD

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    orders.order_id
    , dateadd(day, 7, created_at) as due_date
    , dateadd(month, 1, created_at) as next_month_date
from
    orders
```

期待する和名変換後

```sql
select
    orders.注文ID
    , dateadd(day, 7, 作成日時) as due_date
    , dateadd(month, 1, 作成日時) as next_month_date
from
    orders
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.注文ID |
| 2 | 作成日時 |

### DATEDIFF

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    orders.order_id
    , datediff(day, created_at, updated_at) as elapsed_days
    , datediff(minute, created_at, updated_at) as elapsed_minutes
from
    orders
```

期待する和名変換後

```sql
select
    orders.注文ID
    , datediff(day, 作成日時, 更新日時) as elapsed_days
    , datediff(minute, 作成日時, 更新日時) as elapsed_minutes
from
    orders
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | orders.注文ID |
| 2 | 作成日時 |
| 3 | 更新日時 |

### COUNT

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
    , count(orders.order_id) as order_count
from
    users
    left join orders
        on users.user_id = orders.user_id
group by
    users.user_id
```

期待する和名変換後

```sql
select
    users.ユーザーID
    , count(orders.注文ID) as order_count
from
    users
    left join orders
        on users.ユーザーID = orders.注文ユーザーID
group by
    users.ユーザーID
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | orders.注文ID |
| 3 | orders.注文ユーザーID |

### EXISTS

入力（A5M2 Ctrl+Q 整形後）

```sql
select
    users.user_id
from
    users
where
    exists (
        select
            1
        from
            orders
        where
            orders.user_id = users.user_id
            and orders.amount > 0
    )
```

期待する和名変換後

```sql
select
    users.ユーザーID
from
    users
where
    exists (
        select
            1
        from
            orders
        where
            orders.注文ユーザーID = users.ユーザーID
            and orders.金額 > 0
    )
```

変換内容:

| 順序 | 変換後 |
| --- | --- |
| 1 | users.ユーザーID |
| 2 | orders.注文ユーザーID |
| 3 | orders.金額 |

## 追加T-SQL出力ケース

以下はA5M2 `Ctrl+Q` の実整形を使用した追加ケースです。
確定した期待値は出力期待値ブックへ登録し、手動作成対象は各ケースに明記します。

### SEL-027 JOINのON条件が複数あるケース

入力:

```sql
select
    tb1.user_id
    , tb1.name
    , tb2.order_id
    , tb2.amount
from
    users as tb1
    left join orders as tb2
        on tb1.user_id = tb2.user_id
        and tb2.status = @status
```

変換定義:

| 所属テーブルID | 所属テーブル和名 | カラム名 | カラム和名 |
| --- | --- | --- | --- |
| tb1 | ユーザー | user_id | ユーザーID |
| tb1 | ユーザー | name | 氏名 |
| tb2 | 注文 | order_id | 注文ID |
| tb2 | 注文 | amount | 金額 |
| tb2 | 注文 | user_id | 注文ユーザーID |
| tb2 | 注文 | status | 状態 |

期待値:

- 参照テーブル: `ユーザー[tb1]`、`注文[tb2]`
- 取得項目: `tb1.ユーザーID`、`tb1.氏名`、`tb2.注文ID`、`tb2.金額`
- 結合: `＜ユーザー[tb1] LEFT JOIN 注文[tb2]＞`
- 結合条件1: `tb1.ユーザーID = tb2.注文ユーザーID`
- 結合条件2: `AND tb2.状態 = @status`

### SEL-028 NEXT VALUE FOR

入力:

```sql
select
    next value for dbo.order_sequence as next_order_id
```

変換定義は使用しません。

期待値:

- 参照テーブル: `なし`
- 取得項目: `next_order_id`
- 補足: `NEXT VALUE FOR dbo.order_sequence`

### SEL-029 current_value

入力:

```sql
select
    seq.current_value
from
    sys.sequences as seq
where
    seq.name = 'order_sequence'
```

変換定義:

| 所属テーブルID | 所属テーブル和名 | カラム名 | カラム和名 |
| --- | --- | --- | --- |
| seq | シーケンス | current_value | 現在値 |
| seq | シーケンス | name | シーケンス名 |

期待値:

- 参照テーブル: `シーケンス[seq]`
- 取得項目: `seq.現在値`
- 検索条件: `seq.シーケンス名 = 'order_sequence'`

### SEL-030 OFFSET / FETCH

入力:

```sql
select
    tb1.user_id
    , tb1.name
from
    users as tb1
order by
    tb1.user_id offset 10 rows fetch next 20 rows only
```

期待値:

- 参照テーブル: `ユーザー[tb1]`
- 取得範囲: `OFFSET 10 ROWS FETCH NEXT 20 ROWS ONLY`（取得項目より上の行へ配置し、値はG列から記載）
- 取得項目: `tb1.ユーザーID`、`tb1.氏名`
- 並び順: `tb1.ユーザーID`

### SEL-031 TOP

入力:

```sql
select
    top(10) tb1.user_id
    , tb1.name
from
    users as tb1
order by
    tb1.user_id
```

期待値:

- 参照テーブル: `ユーザー[tb1]`
- 取得件数: `10`（取得項目より上の行へ配置し、値はG列から記載）
- 取得項目: `tb1.ユーザーID`、`tb1.氏名`
- 並び順: `tb1.ユーザーID`

### SEL-032 LIKE

入力:

```sql
select
    tb1.user_id
    , tb1.name
from
    users as tb1
where
    tb1.name like @name + '%'
```

期待値:

- 参照テーブル: `ユーザー[tb1]`
- 取得項目: `tb1.ユーザーID`、`tb1.氏名`
- 検索条件: `tb1.氏名 LIKE @name + '%'`

### SEL-033 UPPER

入力:

```sql
select
    tb1.user_id
    , upper(tb1.name) as upper_name
from
    users as tb1
```

期待値:

- 参照テーブル: `ユーザー[tb1]`
- 取得項目: `tb1.ユーザーID`
- 取得項目: `upper_name`
- 補足: `UPPER(tb1.氏名)`

### SEL-034 LOWER

入力:

```sql
select
    tb1.user_id
    , lower(tb1.name) as lower_name
from
    users as tb1
```

期待値:

- 参照テーブル: `ユーザー[tb1]`
- 取得項目: `tb1.ユーザーID`
- 取得項目: `lower_name`
- 補足: `LOWER(tb1.氏名)`

### SEL-035 CHARINDEX

入力:

```sql
select
    tb1.user_id
    , charindex('@', tb1.email) as at_position
from
    users as tb1
```

期待値:

- 参照テーブル: `ユーザー[tb1]`
- 取得項目: `tb1.ユーザーID`
- 取得項目: `at_position`
- 補足: `CHARINDEX('@', tb1.メール)`

### SEL-036 COLLATE Japanese_BIN

入力:

```sql
select
    tb1.user_id
    , tb1.name
from
    users as tb1
where
    tb1.name collate Japanese_BIN = @name
```

期待値:

- 参照テーブル: `ユーザー[tb1]`
- 取得項目: `tb1.ユーザーID`、`tb1.氏名`
- 検索条件: `tb1.氏名 COLLATE Japanese_BIN = @name`

`COLLATE`は大文字へ統一し、照合順序名 `Japanese_BIN` は変更しません。

### SEL-037 ORDER BY COLLATE Japanese_BIN

入力:

```sql
select
    tb1.user_id
    , tb1.name
from
    users as tb1
order by
    tb1.name collate Japanese_BIN
```

期待値:

- 参照テーブル: `ユーザー[tb1]`
- 取得項目: `tb1.ユーザーID`、`tb1.氏名`
- 並び順: `tb1.氏名 COLLATE Japanese_BIN`

`COLLATE`は大文字へ統一し、照合順序名 `Japanese_BIN` は変更しません。
昇順は既定値のため出力しません。

### SEL-038 GROUP BY

入力:

```sql
select
    tb1.status
    , count(tb1.user_id) as user_count
from
    users as tb1
group by
    tb1.status
```

変換定義:

| 所属テーブルID | 所属テーブル和名 | カラム名 | カラム和名 |
| --- | --- | --- | --- |
| tb1 | ユーザー | status | 状態 |
| tb1 | ユーザー | user_id | ユーザーID |

期待値:

- 参照テーブル: `ユーザー[tb1]`
- 取得項目1: `tb1.状態`
- 取得項目2: `user_count`（補足: `COUNT(tb1.ユーザーID)`）
- グループキー1: `tb1.状態`

以下の未確定ケースは `tests/ManualOutputCases.json` にA5M2の行末空白を含めて保存しています。
開発用ブックへの切り替えには次のコマンドを使用します。

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\Set-ManualOutputCase.ps1 -CaseId SEL-039
```

### SEL-039 HAVING

入力:

```sql
select
    tb1.status
    , count(tb1.user_id) as user_count
from
    users as tb1
group by
    tb1.status
having
    count(tb1.user_id) >= 10
```

変換定義:

| 所属テーブルID | 所属テーブル和名 | カラム名 | カラム和名 |
| --- | --- | --- | --- |
| tb1 | ユーザー | status | 状態 |
| tb1 | ユーザー | user_id | ユーザーID |

期待値:

- 参照テーブル: `ユーザー[tb1]`
- 取得項目1: `tb1.状態`
- 取得項目2: `user_count`（補足: `COUNT(tb1.ユーザーID)`）
- グループキー1: `tb1.状態`
- 集計条件: `COUNT(tb1.ユーザーID) >= 10`

### SEL-040 CASE WHEN ELSE

入力:

```sql
select
    tb1.user_id
    , case
        when tb1.status = 'ACTIVE'
            then '有効'
        else '無効'
        end as status_name
from
    users as tb1
```

変換定義:

| 所属テーブルID | 所属テーブル和名 | カラム名 | カラム和名 |
| --- | --- | --- | --- |
| tb1 | ユーザー | user_id | ユーザーID |
| tb1 | ユーザー | status | 状態 |

期待値:

- 参照テーブル: `ユーザー[tb1]`
- 取得項目1: `tb1.ユーザーID`
- 取得項目2: `status_name`
- CASE式に`AS`エイリアスがある場合は、`CASE結果`ではなくエイリアス名を取得項目へ表示
- CASE分岐1: `tb1.状態 = 'ACTIVE' → '有効'`
- CASE分岐2: `ELSE → '無効'`

### SEL-041 複雑な括弧条件

入力:

```sql
select
    tb1.user_id
    , tb1.name
from
    users as tb1
where
    (
        tb1.status = @status
        and (tb1.name like @name + '%' or tb1.email = @email)
    )
    and not (
        tb1.deleted_at is not null
        or tb1.user_id in (@excluded_user_id1, @excluded_user_id2)
    )
```

変換定義:

| 所属テーブルID | 所属テーブル和名 | カラム名 | カラム和名 |
| --- | --- | --- | --- |
| tb1 | ユーザー | user_id | ユーザーID |
| tb1 | ユーザー | name | 氏名 |
| tb1 | ユーザー | status | 状態 |
| tb1 | ユーザー | email | メール |
| tb1 | ユーザー | deleted_at | 削除日時 |

期待値:

- 参照テーブル: `ユーザー[tb1]`
- 取得項目: `tb1.ユーザーID`、`tb1.氏名`
- 検索条件: `(tb1.状態 = @status AND (tb1.氏名 LIKE @name + '%' OR tb1.メール = @email)) AND NOT (tb1.削除日時 IS NOT NULL OR tb1.ユーザーID IN (@excluded_user_id1, @excluded_user_id2))`

括弧、`AND`、`OR`、`NOT`は検索条件ブロック内の別セルへ配置し、元SQLの評価順序を維持します。

### SEL-042 サブクエリ

入力:

```sql
select
    tb1.user_id
    , tb1.name
from
    users as tb1
where
    tb1.user_id in (
        select
            tb2.user_id
        from
            orders as tb2
        where
            tb2.status = @status
    )
```

変換定義:

| 所属テーブルID | 所属テーブル和名 | カラム名 | カラム和名 |
| --- | --- | --- | --- |
| tb1 | ユーザー | user_id | ユーザーID |
| tb1 | ユーザー | name | 氏名 |
| tb2 | 注文 | user_id | 注文ユーザーID |
| tb2 | 注文 | status | 状態 |

期待値:

- 出力順序: `サブクエリ[SQ1]`、`＜DB入出力項目定義＞`
- SQ1参照テーブル: `注文[tb2]`
- SQ1取得項目: `tb2.注文ユーザーID`
- SQ1検索条件: `tb2.状態 = @status`
- 全体参照テーブル: `ユーザー[tb1]`、`SQ1`
- 全体取得項目: `tb1.ユーザーID`、`tb1.氏名`
- 全体検索条件: `tb1.ユーザーID IN (SQ1)`

### SEL-043 WITH

入力:

```sql
with target_users as (
    select
        tb1.user_id
        , tb1.name
    from
        users as tb1
    where
        tb1.status = @status
)
select
    target_users.user_id
    , target_users.name
from
    target_users
```

変換定義:

| 所属テーブルID | 所属テーブル和名 | カラム名 | カラム和名 |
| --- | --- | --- | --- |
| tb1 | ユーザー | user_id | ユーザーID |
| tb1 | ユーザー | name | 氏名 |
| tb1 | ユーザー | status | 状態 |
| target_users | (和名未取得) | user_id | ユーザーID |
| target_users | (和名未取得) | name | 氏名 |

期待値:

- 出力順序: `サブクエリ[target_users]`、`＜DB入出力項目定義＞`
- WITH内参照テーブル: `ユーザー[tb1]`
- WITH内取得項目: `tb1.ユーザーID`、`tb1.氏名`
- WITH内検索条件: `tb1.状態 = @status`
- 全体参照テーブル: `(和名未取得)[target_users]`
- 全体取得項目: `target_users.ユーザーID`、`target_users.氏名`

CTE名に対応する実テーブル和名は存在しないため、所属テーブル和名を `(和名未取得)` とします。

### SEL-044 UNION

入力:

```sql
select
    tb1.user_id
    , tb1.name
from
    users as tb1
where
    tb1.status = 'ACTIVE'
union
select
    tb2.user_id
    , tb2.name
from
    archived_users as tb2
where
    tb2.status = 'ACTIVE'
```

変換定義:

| 所属テーブルID | 所属テーブル和名 | カラム名 | カラム和名 |
| --- | --- | --- | --- |
| tb1 | ユーザー | user_id | ユーザーID |
| tb1 | ユーザー | name | 氏名 |
| tb1 | ユーザー | status | 状態 |
| tb2 | 退会ユーザー | user_id | ユーザーID |
| tb2 | 退会ユーザー | name | 氏名 |
| tb2 | 退会ユーザー | status | 状態 |

期待値:

- 参照テーブル: `ユーザー[tb1]`、`退会ユーザー[tb2]`
- 前半取得項目: `tb1.ユーザーID`、`tb1.氏名`
- 前半検索条件: `tb1.状態 = 'ACTIVE'`
- UNION境界: `＜UNION＞`
- 後半取得項目: `tb2.ユーザーID`、`tb2.氏名`
- 後半検索条件: `tb2.状態 = 'ACTIVE'`

### SEL-045 INSERT SELECT

入力:

```sql
insert
into user_archive(user_id, name, status)
select
    tb1.user_id
    , tb1.name
    , tb1.status
from
    users as tb1
where
    tb1.deleted_at < @archive_before
```

変換定義:

| 所属テーブルID | 所属テーブル和名 | カラム名 | カラム和名 |
| --- | --- | --- | --- |
| user_archive | ユーザーアーカイブ | user_id | ユーザーID |
| user_archive | ユーザーアーカイブ | name | 氏名 |
| user_archive | ユーザーアーカイブ | status | 状態 |
| - | ユーザーアーカイブ | user_id | ユーザーID |
| - | ユーザーアーカイブ | name | 氏名 |
| - | ユーザーアーカイブ | status | 状態 |
| tb1 | ユーザー | user_id | ユーザーID |
| tb1 | ユーザー | name | 氏名 |
| tb1 | ユーザー | status | 状態 |
| tb1 | ユーザー | deleted_at | 削除日時 |

`user_archive` は移送先テーブル名、`-` は修飾されないINSERT列を変換するための定義です。

期待値:

- 出力順序: `＜DB入出力項目定義＞`、`＜データ移送表＞`
- DB入出力項目定義の参照テーブル: `ユーザー[tb1]`
- DB入出力項目定義の取得項目: `tb1.ユーザーID`、`tb1.氏名`、`tb1.状態`
- DB入出力項目定義の検索条件: `tb1.削除日時 < @archive_before`
- データ移送表の参照テーブル: `ユーザーアーカイブ`
- 表見出し: `項目`、`移送元`、`移送方法ほか`
- 移送対応: `ユーザーID ← tb1.ユーザーID`
- 移送対応: `氏名 ← tb1.氏名`
- 移送対応: `状態 ← tb1.状態`
- 移送方法ほか: 空欄

### SEL-046 DELETE

入力:

```sql
delete tb1
from
    users as tb1
where
    tb1.deleted_at < @delete_before
    and tb1.status = 'INACTIVE'
```

変換定義:

| 所属テーブルID | 所属テーブル和名 | カラム名 | カラム和名 |
| --- | --- | --- | --- |
| tb1 | ユーザー | deleted_at | 削除日時 |
| tb1 | ユーザー | status | 状態 |

期待値:

- タイトル: `＜データ移送表＞`
- 参照テーブル: `ユーザー[tb1]`
- 検索条件1: `tb1.削除日時 < @delete_before`
- 論理演算子: `AND`
- 検索条件2: `tb1.状態 = 'INACTIVE'`
- DELETEでは移送項目を出力しない

### SEL-047 UPDATE FROM

入力:

```sql
update tb1
set
    status = @status
    , updated_at = sysdatetime()
from
    users as tb1
    inner join orders as tb2
        on tb1.user_id = tb2.user_id
where
    tb2.status = @order_status
```

変換定義:

| 所属テーブルID | 所属テーブル和名 | カラム名 | カラム和名 |
| --- | --- | --- | --- |
| - | ユーザー | status | 状態 |
| - | ユーザー | updated_at | 更新日時 |
| tb1 | ユーザー | user_id | ユーザーID |
| tb2 | 注文 | user_id | 注文ユーザーID |
| tb2 | 注文 | status | 状態 |

期待値:

- タイトル: `＜データ移送表＞`
- 参照テーブル: `ユーザー[tb1]`、`注文[tb2]`
- 表見出し: `項目`、`移送元`、`移送方法ほか`
- 移送項目: `状態`、移送元は空欄、移送方法ほかは`@status`
- 移送項目: `更新日時`、移送元は空欄、移送方法ほかは`sysdatetime()`
- 結合対象: `＜ユーザー[tb1] INNER JOIN 注文[tb2]＞`
- 結合条件: `tb1.ユーザーID = tb2.注文ユーザーID`
- 検索条件: `tb2.状態 = @order_status`

## 出力シート

整形出力先シート名は `アウトプット` です。
SELECT系は`＜DB入出力項目定義＞`、更新系は`＜データ移送表＞`としてA列からCL列へ出力します。
WITH句やサブクエリを含む場合は、内側のサブクエリ、外側のサブクエリ、解析対象のクエリ全体の順で出力します。
`FROM (VALUES ...) AS 別名`は、参照テーブルへ`派生テーブル[別名]`と表示します。
ネストしたCASEは親条件と内側の先頭条件をAF列で直接連結し、内側の残りの分岐を2列右のAH列から表示します。
SELECTとGROUP BYで同じCASE式を使う場合は、SELECT側のエイリアス名をグループキーへ表示し、CASE分岐も再掲します。
ORDER BY内のCASEにエイリアスがない場合は、ソートキー名を`CASE結果`としてCASE分岐を右側へ表示します。
関数や演算式に含まれるCASEは、外側の式を`SUM(CASE結果)`のように残します。単一CASEの分岐は外側式の開始位置から8列右へ下げます。1式に複数のCASEがある場合は、番号付きの`CASE結果n`を外側式から14列下げ、各分岐を`CASE結果n`から6列下げます。
列エイリアスのないCASEを取得項目として直接返す場合は、Q列へ`CASE結果`、AF列以降へ分岐を複数行で表示します。
複合WHEN条件はCASE直下の先頭条件とAND/ORを基準列へ置き、後続条件を2列右へ下げます。THENの結果は最後の条件と同じ列、ELSEは原文どおり`ELSE`として基準列へ表示します。THENまたはELSEの直下にCASEがある場合は親条件と内側の先頭条件を同じ行へ直接連結し、内側の残りの分岐をさらに一段右へ表示します。
ORDER BY内のIIFは分岐へ展開せず、式全体をソートキーとして表示します。
複雑な条件式は括弧の階層を2列ずつ右へ展開し、各階層の論理演算子と次の条件を別セルに表示します。
SELECT DISTINCTは取得項目より上の`重複除外`行へ`DISTINCT`と表示します。
SELECT INTOとINSERT SELECTのトップレベルSELECTは人工的なサブクエリにせず、`＜DB入出力項目定義＞`の取得式を`＜データ移送表＞`へ直接対応させます。SELECT INTO、INSERT、UPDATE、DELETEのデータ移送表では、参照テーブルを移送先、移送元の順に`、`区切りで表示します。移送元を持たない通常のINSERT VALUESは移送先だけを表示します。実在するサブクエリ、CTE、派生テーブルだけは先にサブクエリ表として表示し、INSERT VALUESの値式にあるスカラーサブクエリも同様にSQ名で参照します。
SELECT取得式の未修飾列は、FROM句と変換定義から所属テーブルが一意に決まる場合だけSQL上の別名を補い、取得項目とデータ移送表の移送元へ反映します。A列が`-`の場合も、同じB列を持つA列付きの別定義を介して1つのSQL別名へ一意に対応すればその別名を補い、D列に和名があれば和名、空欄なら物理列名を表示します。複数テーブルに一致する場合は未修飾のまま保持します。
SQL解析シートのC列以降へ出力する変換内容は、アウトプット解析で同じ未修飾列の所属先が一意に決まった場合、`名前`から`tb1.名前`のように同じプレフィックス付き表示へ更新します。変換内容とAST補完結果はメモリ上で統合して最後に一度だけ書き込み、同じ最終値は重複表示しません。変換内容がない物理列をC列へ新規追加はしません。
CASE式の列エイリアスは、WHEN条件を除外し、ネストを含む全THEN・ELSE末端が単一列由来で、その物理列名と和名がすべて一致し、エイリアスも同じ物理列名の場合だけ和名へ解決します。TRIMやCASTで包まれた単一列は列由来として扱います。ELSEなし、定数結果、複数列結果、物理列名または和名の不一致、異なるエイリアスでは元のエイリアスを保持します。
INSERT SELECTのトップレベルにUNIONまたはUNION ALLがある場合、SELECT側の集合演算表示は維持し、データ移送表側では左から順に`＜移送パターン1＞`、`＜移送パターン2＞`として各取得式をINSERT対象列へ対応させます。
複数行のINSERT VALUESは、SQLに記述された順に`＜VALUES 1行目＞`、`＜VALUES 2行目＞`として独立したデータ移送表へ分け、各ラベルは表本体の外に置きます。
未対応のクエリは和名変換後SQLをA列へ1行ずつ出力し、1行空けてフォールバック原因を表示します。

## ユーザー確認済みの追加期待値ケース

次のSQLと和名定義を`tests/ManualOutputCases.json`へ登録し、ユーザーレビューで確定した期待値を`tests/SqlAnalysisFormatter.OutputExpectations.xlsx`へ保存しています。

| ケース | 内容 |
| --- | --- |
| SEL-048 | SELECT FROM VALUES |
| SEL-049 | ネストしたCASE |
| SEL-050 | GROUP BY CASE |
| SEL-051 | ORDER BY CASE |
| SEL-052 | ORDER BY IIF |
| SEL-053 | 多段の複雑な条件式 |
| SEL-054 | 複雑な派生テーブルJOIN |
| SEL-055 | ネストしたサブクエリ |
| SEL-057 | WHERE句の複数WHENで各先頭条件を基準列へ揃える検索CASE |
| SEL-058 | SELECT DISTINCT |
| SEL-059 | SELECT INTOのトップレベルSELECT直接対応 |
| SEL-061 | CASEを含む複雑なSELECT INTOのトップレベルSELECT直接対応 |
| SEL-062 | UPDATE SETの実スカラーサブクエリをSQ1として先に出し、移送元へ対応 |
| SEL-063 | 派生SELECTをサブクエリ表として先に出し、UPDATE移送元へ派生列を対応 |
| SEL-064 | DELETEのEXISTS内SELECTをSQ1として分離し、削除条件をEXISTS (SQ1)へ置換 |
| SEL-065 | FROMなしUPDATEのEXISTS内SELECTをSQ1として分離し、更新対象とSQ1を参照表示 |
| SEL-066 | UPDATE比較条件の集計SELECTをSQ1として分離し、比較対象を(SQ1)へ置換 |
| SEL-067 | UPDATEのIN句内SELECTをSQ1として分離し、外側検索条件をIN (SQ1)へ置換 |
| SEL-068 | SUM(CASE結果)の分岐を外側式の開始位置から8列下げて表示 |
| SEL-069 | ELSE内のネストCASEを親ELSEへ直接連結し、ELSE表記を原文どおり出力 |
| SEL-072 | INSERT VALUESの列値を返さないCASEを移送方法へ配置し、同一項目の複数行を1つの枠で表示 |
| SEL-045 | 単純なINSERT SELECTのトップレベルSELECT直接対応 |
| SEL-056 | 複雑なINSERT SELECTのトップレベルSELECT直接対応 |
| SEL-070 | 複合WHENの先頭条件を基準列、ANDの後続条件を2列下げて階層表示 |
| SEL-071 | UPDATE SETのCASEを移送方法へ配置し、先頭条件とANDを同じ基準列で表示 |
| SEL-060 | INSERT VALUESの変数・定数・関数を移送方法へ配置し、テーブル列を参照しない移送元を空欄表示 |
| SEL-076 | 複数行INSERT VALUESを行ごとの独立したデータ移送表へ分けて表示 |
| SEL-077 | INSERT SELECT内のUNION ALLをSELECT表と分岐別の移送パターンへ分けて表示 |
| SEL-078 | 1式内の複数CASEを番号付きの結果へ分け、外側式から14列、分岐を結果から6列下げて表示 |
| SEL-079 | 別々のCASE取得項目に`paid_amount`、`refund_amount`のエイリアスを表示 |
| SEL-080 | 集計関数で包まれたCASE取得項目ごとにエイリアスと8列字下げした分岐を表示 |
| SEL-081 | THEN・ELSE両側を3段までネストし、親分岐と各CASEの先頭分岐を直接連結して表示 |
| SEL-082 | 複合条件を持つTHEN・ELSE両側のネストCASEを、AND・ORと括弧階層を保って独立行へ表示 |
| SEL-075 | 取得結果を直接返すCASEを`CASE結果`として分岐とともに表示 |
| SEL-073 | TOP内のCASEを`取得件数`の`CASE結果`として分岐とともに表示 |
| SEL-074 | OFFSET内のCASEを大文字の取得範囲へ置換し、分岐をAA/AB列へ表示 |

SEL-080までの各SQLはScriptDomで構文エラーがないことと、A5M2 `Ctrl+Q`の実整形結果を確認済みです。SEL-081とSEL-082はScriptDomで解析できること出力配置をユーザーレビューで確定済みです。

## 期待値レビュー待ち

現在、期待値レビュー待ちのケースはありません。

## 暫定実装済み・ユーザーレビュー待ち

現在、ユーザーレビュー待ちの暫定実装ケースはありません。

SEL-060、SEL-062からSEL-080はA5M2 2.21.2の`Ctrl+Q`形式に合わせた入力を使用しています。
