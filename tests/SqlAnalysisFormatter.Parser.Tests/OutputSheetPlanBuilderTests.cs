using SqlAnalysisFormatter.Parser;

namespace SqlAnalysisFormatter.Parser.Tests;

/// <summary>
/// ASTからアウトプット描画計画を作る仕様テスト
/// </summary>
[TestClass]
public sealed class OutputSheetPlanBuilderTests
{
    /// <summary>
    /// 単純SELECTを基本フレームへ変換できることを確認
    /// </summary>
    [TestMethod]
    public void Build_CreatesSimpleSelectFrame()
    {
        const string sql = """
            SELECT
                tb1.ユーザーID
                , tb1.氏名
            FROM
                users AS tb1
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "user_id", "ユーザーID")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual(4, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "＜DB入出力項目定義＞"),
            (2, 1, "参照テーブル: ユーザー[tb1]"),
            (3, 1, "取得項目"),
            (3, 7, "取得項目1"),
            (3, 15, ":"),
            (3, 17, "tb1.ユーザーID"),
            (4, 7, "取得項目2"),
            (4, 15, ":"),
            (4, 17, "tb1.氏名"));
        CollectionAssert.AreEqual(
            new[]
            {
                new OutputSection(OutputSectionKind.Reference, 2, 2),
                new OutputSection(OutputSectionKind.Standard, 3, 4)
            },
            plan.Sections.ToArray());
    }

    /// <summary>
    /// 単一テーブルでは単独フィールド定義のテーブル和名を参照表示へ使用することを確認
    /// </summary>
    [TestMethod]
    public void Build_UsesStandaloneMappingTableNameForSingleTable()
    {
        const string sql = "select name from [user]";
        MappingDefinition[] mappings =
        [
            new("-", "ユーザー", "name", "名前")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual("参照テーブル: ユーザー", CellValue(plan, 2, 1));
    }

    /// <summary>
    /// 単独フィールド定義から解決したテーブル和名にも明示別名を付けることを確認
    /// </summary>
    [TestMethod]
    public void Build_KeepsAliasForStandaloneMappingTableName()
    {
        const string sql = "select u.name from [user] as u";
        MappingDefinition[] mappings =
        [
            new("-", "ユーザー", "name", "名前")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual("参照テーブル: ユーザー[u]", CellValue(plan, 2, 1));
    }

    /// <summary>
    /// 単独フィールド定義のテーブル和名が複数ある場合は推測しないことを確認
    /// </summary>
    [TestMethod]
    public void Build_DoesNotGuessAmbiguousStandaloneMappingTableName()
    {
        const string sql = "select name from [user]";
        MappingDefinition[] mappings =
        [
            new("-", "ユーザー", "name", "名前"),
            new("-", "注文", "order_id", "注文ID")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual("参照テーブル: (和名未取得)[user]", CellValue(plan, 2, 1));
    }

    /// <summary>
    /// parser用識別子を構文文字を含む和名へ復元できることを確認
    /// </summary>
    [TestMethod]
    public void Build_RestoresParserFieldIdentifiersContainingSqlSyntax()
    {
        const string sql = """
            SELECT
                tb1.__SAF_FIELD_R000002__
                , __SAF_FIELD_R000003__
                , tb1.__SAF_FIELD_R000004__
            FROM
                invoices AS tb1
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "請求", "amount", "請求額[税込]/月", "__SAF_FIELD_R000002__"),
            new("-", "", "status", "状態 - 判定(仮), 100%", "__SAF_FIELD_R000003__"),
            new("tb1", "請求", "owner", "担当者's/*主*/--現行", "__SAF_FIELD_R000004__")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual("tb1.請求額[税込]/月", CellValue(plan, 3, 17));
        Assert.AreEqual("状態 - 判定(仮), 100%", CellValue(plan, 4, 17));
        Assert.AreEqual("tb1.担当者's/*主*/--現行", CellValue(plan, 5, 17));
        Assert.IsFalse(plan.Cells.Any(cell => cell.Value.Contains("__SAF_FIELD_", StringComparison.Ordinal)));
    }

    /// <summary>
    /// 復元した和名を別のparser用IDとして再変換しないことを確認
    /// </summary>
    [TestMethod]
    public void Build_DoesNotRestoreParserFieldIdentifiersRecursively()
    {
        const string sql = "SELECT tb1.__SAF_FIELD_R000002__, tb1.__SAF_FIELD_R000003__ FROM users AS tb1";
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "first", "参照__SAF_FIELD_R000003__", "__SAF_FIELD_R000002__"),
            new("tb1", "ユーザー", "second", "氏名", "__SAF_FIELD_R000003__")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual("tb1.参照__SAF_FIELD_R000003__", CellValue(plan, 3, 17));
        Assert.AreEqual("tb1.氏名", CellValue(plan, 4, 17));
    }

    /// <summary>
    /// COALESCEの関数名だけを大文字へ統一することを確認
    /// </summary>
    [TestMethod]
    public void Build_NormalizesCoalesceKeywordToUppercase()
    {
        const string sql = """
            SELECT
                coalesce(tb1.メール, coalesce(tb1.氏名, 'coalesce(')) AS contact_text
            FROM
                users AS tb1
            ORDER BY
                coalesce(tb1.メール, tb1.氏名)
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "email", "メール")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual(
            "COALESCE(tb1.メール, COALESCE(tb1.氏名, 'coalesce('))",
            CellValue(plan, 3, 32));
        Assert.AreEqual(
            "COALESCE(tb1.メール, tb1.氏名)",
            CellValue(plan, 4, 17));
    }

    /// <summary>
    /// 組み込み関数とSQL演算子を帳票向けの大文字表記へ統一することを確認
    /// </summary>
    [TestMethod]
    public void Build_NormalizesFunctionsAndSqlOperatorsForDisplay()
    {
        const string sql = """
            select
                count(tb1.ユーザーID) as user_count
                , sysdatetime() as created_at
            from
                users as tb1
            order by
                iif(tb1.削除日時 is null, 0, 1)
            """;

        var plan = OutputSheetPlanBuilder.Build(sql, [new("tb1", "ユーザー", "", "")]);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual("COUNT(tb1.ユーザーID)", CellValue(plan, 3, 32));
        Assert.AreEqual("SYSDATETIME()", CellValue(plan, 4, 32));
        Assert.AreEqual("IIF(tb1.削除日時 IS NULL, 0, 1)", CellValue(plan, 5, 17));
    }

    /// <summary>
    /// VALUESテーブル値コンストラクターを派生テーブルとして表示することを確認
    /// </summary>
    [TestMethod]
    public void Build_WritesValuesSourceAsDerivedTable()
    {
        const string sql = """
            select
                v.ユーザーID
            from
                (values (1), (2)) as v(user_id)
            """;

        var plan = OutputSheetPlanBuilder.Build(sql, []);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual("参照テーブル: 派生テーブル[v]", CellValue(plan, 2, 1));
    }

    /// <summary>
    /// DISTINCT指定を取得項目より前へ出力することを確認
    /// </summary>
    [TestMethod]
    public void Build_WritesDistinctBeforeSelectItems()
    {
        const string sql = """
            select distinct
                tb1.状態
            from
                users as tb1
            """;

        var plan = OutputSheetPlanBuilder.Build(sql, [new("tb1", "ユーザー", "", "")]);

        Assert.AreEqual(4, plan.RowCount);
        Assert.AreEqual("重複除外", CellValue(plan, 3, 1));
        Assert.AreEqual("DISTINCT", CellValue(plan, 3, 7));
        Assert.AreEqual("取得項目", CellValue(plan, 4, 1));
    }

    /// <summary>
    /// SELECT各句を期待する行順で出力することを確認
    /// </summary>
    [TestMethod]
    public void Build_CreatesSelectClausesInOutputOrder()
    {
        const string sql = """
            SELECT TOP (10)
                tb1.状態
                , COUNT(tb1.ユーザーID) AS user_count
            FROM
                users AS tb1
            WHERE
                tb1.状態 = 'ACTIVE'
            GROUP BY
                tb1.状態
            HAVING
                COUNT(tb1.ユーザーID) >= 10
            ORDER BY
                tb1.状態 DESC
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "status", "状態")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual(9, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "＜DB入出力項目定義＞"),
            (2, 1, "参照テーブル: ユーザー[tb1]"),
            (3, 1, "取得件数"),
            (3, 7, "10"),
            (4, 1, "取得項目"),
            (4, 7, "取得項目1"),
            (4, 15, ":"),
            (4, 17, "tb1.状態"),
            (5, 7, "取得項目2"),
            (5, 15, ":"),
            (5, 17, "user_count"),
            (5, 31, "※"),
            (5, 32, "COUNT(tb1.ユーザーID)"),
            (6, 1, "検索条件"),
            (6, 17, "tb1.状態 = 'ACTIVE'"),
            (7, 1, "グループ"),
            (7, 7, "グループキー1"),
            (7, 15, ":"),
            (7, 17, "tb1.状態"),
            (8, 1, "集計条件"),
            (8, 17, "COUNT(tb1.ユーザーID) >= 10"),
            (9, 1, "並び順"),
            (9, 7, "ソートキー1"),
            (9, 15, ":"),
            (9, 17, "tb1.状態(降順)"));
    }

    /// <summary>
    /// OFFSETとFETCHを取得項目より先に出力することを確認
    /// </summary>
    [TestMethod]
    public void Build_WritesOffsetFetchBeforeSelectItems()
    {
        const string sql = """
            SELECT
                tb1.ユーザーID
                , tb1.氏名
            FROM
                users AS tb1
            ORDER BY
                tb1.ユーザーID
            OFFSET 10 ROWS FETCH NEXT 20 ROWS ONLY
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "user_id", "ユーザーID")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.AreEqual(6, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "＜DB入出力項目定義＞"),
            (2, 1, "参照テーブル: ユーザー[tb1]"),
            (3, 1, "取得範囲"),
            (3, 7, "OFFSET 10 ROWS FETCH NEXT 20 ROWS ONLY"),
            (4, 1, "取得項目"),
            (4, 7, "取得項目1"),
            (4, 15, ":"),
            (4, 17, "tb1.ユーザーID"),
            (5, 7, "取得項目2"),
            (5, 15, ":"),
            (5, 17, "tb1.氏名"),
            (6, 1, "並び順"),
            (6, 7, "ソートキー1"),
            (6, 15, ":"),
            (6, 17, "tb1.ユーザーID"));
    }

    /// <summary>
    /// 複数ON条件を持つJOINフレームを作成できることを確認
    /// </summary>
    [TestMethod]
    public void Build_CreatesJoinFrameWithMultipleConditions()
    {
        const string sql = """
            SELECT
                tb1.ユーザーID
                , tb1.氏名
                , tb2.注文ID
                , tb2.金額
            FROM
                users AS tb1
                LEFT JOIN orders AS tb2
                    ON tb1.ユーザーID = tb2.注文ユーザーID
                    AND tb2.状態 = @status
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "user_id", "ユーザーID"),
            new("tb2", "注文", "order_id", "注文ID")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual(9, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "＜DB入出力項目定義＞"),
            (2, 1, "参照テーブル: ユーザー[tb1]、注文[tb2]"),
            (3, 1, "取得項目"),
            (3, 7, "取得項目1"),
            (3, 15, ":"),
            (3, 17, "tb1.ユーザーID"),
            (4, 7, "取得項目2"),
            (4, 15, ":"),
            (4, 17, "tb1.氏名"),
            (5, 7, "取得項目3"),
            (5, 15, ":"),
            (5, 17, "tb2.注文ID"),
            (6, 7, "取得項目4"),
            (6, 15, ":"),
            (6, 17, "tb2.金額"),
            (7, 1, "結合条件"),
            (7, 17, "＜ユーザー[tb1] LEFT JOIN 注文[tb2]＞"),
            (8, 17, "tb1.ユーザーID = tb2.注文ユーザーID"),
            (9, 7, "AND"),
            (9, 17, "tb2.状態 = @status"));
    }

    /// <summary>
    /// CASE分岐を列エイリアスの下へ展開することを確認
    /// </summary>
    [TestMethod]
    public void Build_ExpandsCaseBranchesBelowAliasedItem()
    {
        const string sql = """
            SELECT
                tb1.ユーザーID
                , CASE
                    WHEN tb1.状態 = 'ACTIVE'
                        THEN '有効'
                    ELSE '無効'
                    END AS status_name
            FROM
                users AS tb1
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "status", "状態")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.AreEqual(5, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "＜DB入出力項目定義＞"),
            (2, 1, "参照テーブル: ユーザー[tb1]"),
            (3, 1, "取得項目"),
            (3, 7, "取得項目1"),
            (3, 15, ":"),
            (3, 17, "tb1.ユーザーID"),
            (4, 7, "取得項目2"),
            (4, 15, ":"),
            (4, 17, "status_name"),
            (4, 31, "※"),
            (4, 32, "tb1.状態 = 'ACTIVE' → '有効'"),
            (5, 32, "それ以外 → '無効'"));
    }

    /// <summary>
    /// ネストしたCASEを階層ごとの列へ展開することを確認
    /// </summary>
    [TestMethod]
    public void Build_ExpandsNestedCaseBranches()
    {
        const string sql = """
            select
                case
                    when tb1.状態 = 'ACTIVE' then
                        case
                            when tb1.ランクコード = 'VIP' then '優良'
                            else '通常'
                        end
                    when tb1.状態 = 'LOCKED' then '停止'
                    else '無効'
                end as user_category
            from
                users as tb1
            """;

        var plan = OutputSheetPlanBuilder.Build(sql, [new("tb1", "ユーザー", "", "")]);

        Assert.AreEqual(7, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "＜DB入出力項目定義＞"),
            (2, 1, "参照テーブル: ユーザー[tb1]"),
            (3, 1, "取得項目"),
            (3, 7, "取得項目1"),
            (3, 15, ":"),
            (3, 17, "user_category"),
            (3, 31, "※"),
            (3, 32, "tb1.状態 = 'ACTIVE' → CASE"),
            (4, 34, "tb1.ランクコード = 'VIP' → '優良'"),
            (5, 34, "それ以外 → '通常'"),
            (6, 32, "tb1.状態 = 'LOCKED' → '停止'"),
            (7, 32, "それ以外 → '無効'"));
    }

    /// <summary>
    /// GROUP BYのCASEを取得項目のエイリアスと分岐へ展開することを確認
    /// </summary>
    [TestMethod]
    public void Build_ExpandsCaseUsedByGroupBy()
    {
        const string sql = """
            select
                case
                    when tb1.金額 >= 10000 then 'HIGH'
                    else 'NORMAL'
                end as amount_band
            from
                orders as tb1
            group by
                case
                    when tb1.金額 >= 10000 then 'HIGH'
                    else 'NORMAL'
                end
            """;

        var plan = OutputSheetPlanBuilder.Build(sql, [new("tb1", "注文", "", "")]);

        Assert.AreEqual(6, plan.RowCount);
        Assert.AreEqual("amount_band", CellValue(plan, 5, 17));
        Assert.AreEqual("※", CellValue(plan, 5, 31));
        Assert.AreEqual("tb1.金額 >= 10000 → 'HIGH'", CellValue(plan, 5, 32));
        Assert.AreEqual("それ以外 → 'NORMAL'", CellValue(plan, 6, 32));
    }

    /// <summary>
    /// ORDER BYのCASEを分岐へ展開して後続キーを次の行へ送ることを確認
    /// </summary>
    [TestMethod]
    public void Build_ExpandsCaseUsedByOrderBy()
    {
        const string sql = """
            select
                tb1.ユーザーID
            from
                users as tb1
            order by
                case
                    when tb1.状態 = 'ACTIVE' then 0
                    else 1
                end
                , tb1.ユーザーID
            """;

        var plan = OutputSheetPlanBuilder.Build(sql, [new("tb1", "ユーザー", "", "")]);

        Assert.AreEqual(6, plan.RowCount);
        Assert.AreEqual("CASE結果", CellValue(plan, 4, 17));
        Assert.AreEqual("tb1.状態 = 'ACTIVE' → 0", CellValue(plan, 4, 32));
        Assert.AreEqual("それ以外 → 1", CellValue(plan, 5, 32));
        Assert.AreEqual("ソートキー2", CellValue(plan, 6, 7));
        Assert.AreEqual("tb1.ユーザーID", CellValue(plan, 6, 17));
    }

    /// <summary>
    /// WHEREで比較するCASEを比較結果と分岐へ展開することを確認
    /// </summary>
    [TestMethod]
    public void Build_ExpandsCaseComparedByWhere()
    {
        const string sql = """
            select
                tb1.ユーザーID
            from
                users as tb1
            where
                case
                    when tb1.状態 = 'ACTIVE' and tb1.削除日時 is null then 1
                    else 0
                end = 1
            """;

        var plan = OutputSheetPlanBuilder.Build(sql, [new("tb1", "ユーザー", "", "")]);

        Assert.AreEqual(5, plan.RowCount);
        Assert.AreEqual("CASE結果 = 1", CellValue(plan, 4, 17));
        Assert.AreEqual("※", CellValue(plan, 4, 31));
        Assert.AreEqual("tb1.状態 = 'ACTIVE' AND tb1.削除日時 IS NULL → 1", CellValue(plan, 4, 32));
        Assert.AreEqual("それ以外 → 0", CellValue(plan, 5, 32));
    }

    /// <summary>
    /// ORで分岐する深い括弧条件を列階層へ展開することを確認
    /// </summary>
    [TestMethod]
    public void Build_ExpandsDeepBooleanGroups()
    {
        const string sql = """
            select
                tb1.ユーザーID
            from
                users as tb1
            where
                (
                    (
                        tb1.状態 in ('ACTIVE', 'LOCKED')
                        and (tb1.メール like @domain or tb1.氏名 like @name)
                    )
                    or (
                        tb1.状態 = 'PENDING'
                        and tb1.削除日時 is null
                    )
                )
                and not (
                    tb1.ユーザーID between @from and @to
                    or coalesce(tb1.上司ID, 0) = @manager_id
                )
            """;

        var plan = OutputSheetPlanBuilder.Build(sql, [new("tb1", "ユーザー", "", "")]);

        Assert.AreEqual(13, plan.RowCount);
        Assert.AreEqual("(", CellValue(plan, 4, 7));
        Assert.AreEqual("(", CellValue(plan, 4, 15));
        Assert.AreEqual("tb1.状態 IN ('ACTIVE', 'LOCKED')", CellValue(plan, 4, 17));
        Assert.AreEqual("AND", CellValue(plan, 5, 17));
        Assert.AreEqual("(tb1.メール LIKE @domain OR tb1.氏名 LIKE @name)", CellValue(plan, 5, 19));
        Assert.AreEqual("OR", CellValue(plan, 7, 15));
        Assert.AreEqual("(", CellValue(plan, 7, 17));
        Assert.AreEqual("AND NOT", CellValue(plan, 11, 7));
        Assert.AreEqual("COALESCE(tb1.上司ID, 0) = @manager_id", CellValue(plan, 12, 17));
    }

    /// <summary>
    /// 日付計算の符号前後に入った空白を帳票上で正規化することを確認
    /// </summary>
    [TestMethod]
    public void Build_NormalizesSpacedUnarySignInCondition()
    {
        const string sql = """
            select
                tb1.ユーザーID
            from
                users as tb1
            where
                tb1.作成日時 >= dateadd(day, - 30, @base_date)
            """;

        var plan = OutputSheetPlanBuilder.Build(sql, [new("tb1", "ユーザー", "", "")]);

        Assert.AreEqual("tb1.作成日時 >= DATEADD(DAY, -30, @base_date)", CellValue(plan, 4, 17));
    }

    /// <summary>
    /// 複雑条件の外側括弧だけを独立行へ展開することを確認
    /// </summary>
    [TestMethod]
    public void Build_ExpandsOnlyOuterConditionParentheses()
    {
        const string sql = """
            SELECT
                tb1.ユーザーID
                , tb1.氏名
            FROM
                users AS tb1
            WHERE
                (
                    tb1.状態 = @status
                    AND (tb1.氏名 LIKE @name + '%' OR tb1.メール = @email)
                )
                AND NOT (
                    tb1.削除日時 IS NOT NULL
                    OR tb1.ユーザーID IN (@excluded_user_id1, @excluded_user_id2)
                )
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "user_id", "ユーザーID")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.AreEqual(10, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "＜DB入出力項目定義＞"),
            (2, 1, "参照テーブル: ユーザー[tb1]"),
            (3, 1, "取得項目"),
            (3, 7, "取得項目1"),
            (3, 15, ":"),
            (3, 17, "tb1.ユーザーID"),
            (4, 7, "取得項目2"),
            (4, 15, ":"),
            (4, 17, "tb1.氏名"),
            (5, 1, "検索条件"),
            (5, 7, "("),
            (5, 17, "tb1.状態 = @status"),
            (6, 15, "AND"),
            (6, 17, "(tb1.氏名 LIKE @name + '%' OR tb1.メール = @email)"),
            (7, 7, ")"),
            (8, 7, "AND NOT"),
            (8, 15, "("),
            (8, 17, "tb1.削除日時 IS NOT NULL"),
            (9, 15, "OR"),
            (9, 17, "tb1.ユーザーID IN (@excluded_user_id1, @excluded_user_id2)"),
            (10, 7, ")"));
    }

    /// <summary>
    /// 無名サブクエリを全体クエリより先に出力することを確認
    /// </summary>
    [TestMethod]
    public void Build_WritesUnnamedSubqueryBeforeWholeQuery()
    {
        const string sql = """
            SELECT
                tb1.ユーザーID
                , tb1.氏名
            FROM
                users AS tb1
            WHERE
                tb1.ユーザーID IN (
                    SELECT
                        tb2.注文ユーザーID
                    FROM
                        orders AS tb2
                    WHERE
                        tb2.状態 = @status
                )
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "user_id", "ユーザーID"),
            new("tb2", "注文", "user_id", "注文ユーザーID")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.AreEqual(10, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "サブクエリ[SQ1]"),
            (2, 1, "参照テーブル: 注文[tb2]"),
            (3, 1, "取得項目"),
            (3, 7, "取得項目1"),
            (3, 15, ":"),
            (3, 17, "tb2.注文ユーザーID"),
            (4, 1, "検索条件"),
            (4, 17, "tb2.状態 = @status"),
            (6, 1, "＜DB入出力項目定義＞"),
            (7, 1, "参照テーブル: ユーザー[tb1]、SQ1"),
            (8, 1, "取得項目"),
            (8, 7, "取得項目1"),
            (8, 15, ":"),
            (8, 17, "tb1.ユーザーID"),
            (9, 7, "取得項目2"),
            (9, 15, ":"),
            (9, 17, "tb1.氏名"),
            (10, 1, "検索条件"),
            (10, 17, "tb1.ユーザーID IN (SQ1)"));
    }

    /// <summary>
    /// CTE名を使ったサブクエリフレームを先に出力することを確認
    /// </summary>
    [TestMethod]
    public void Build_WritesCteBeforeWholeQueryUsingItsName()
    {
        const string sql = """
            WITH target_users AS (
                SELECT
                    tb1.ユーザーID
                    , tb1.氏名
                FROM
                    users AS tb1
                WHERE
                    tb1.状態 = @status
            )
            SELECT
                target_users.ユーザーID
                , target_users.氏名
            FROM
                target_users
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "user_id", "ユーザーID"),
            new("target_users", "(和名未取得)", "user_id", "ユーザーID")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.AreEqual(10, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "サブクエリ[target_users]"),
            (2, 1, "参照テーブル: ユーザー[tb1]"),
            (3, 1, "取得項目"),
            (3, 7, "取得項目1"),
            (3, 15, ":"),
            (3, 17, "tb1.ユーザーID"),
            (4, 7, "取得項目2"),
            (4, 15, ":"),
            (4, 17, "tb1.氏名"),
            (5, 1, "検索条件"),
            (5, 17, "tb1.状態 = @status"),
            (7, 1, "＜DB入出力項目定義＞"),
            (8, 1, "参照テーブル: (和名未取得)[target_users]"),
            (9, 1, "取得項目"),
            (9, 7, "取得項目1"),
            (9, 15, ":"),
            (9, 17, "target_users.ユーザーID"),
            (10, 7, "取得項目2"),
            (10, 15, ":"),
            (10, 17, "target_users.氏名"));
    }

    /// <summary>
    /// UNIONの各枝を1フレームへ出力することを確認
    /// </summary>
    [TestMethod]
    public void Build_WritesUnionBranchesInOneFrame()
    {
        const string sql = """
            SELECT
                tb1.ユーザーID
                , tb1.氏名
            FROM
                users AS tb1
            WHERE
                tb1.状態 = 'ACTIVE'
            UNION
            SELECT
                tb2.ユーザーID
                , tb2.氏名
            FROM
                archived_users AS tb2
            WHERE
                tb2.状態 = 'ACTIVE'
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "user_id", "ユーザーID"),
            new("tb2", "退会ユーザー", "user_id", "ユーザーID")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.AreEqual(9, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "＜DB入出力項目定義＞"),
            (2, 1, "参照テーブル: ユーザー[tb1]、退会ユーザー[tb2]"),
            (3, 1, "取得項目"),
            (3, 7, "取得項目1"),
            (3, 15, ":"),
            (3, 17, "tb1.ユーザーID"),
            (4, 7, "取得項目2"),
            (4, 15, ":"),
            (4, 17, "tb1.氏名"),
            (5, 1, "検索条件"),
            (5, 17, "tb1.状態 = 'ACTIVE'"),
            (6, 1, "＜UNION＞"),
            (7, 1, "取得項目"),
            (7, 7, "取得項目1"),
            (7, 15, ":"),
            (7, 17, "tb2.ユーザーID"),
            (8, 7, "取得項目2"),
            (8, 15, ":"),
            (8, 17, "tb2.氏名"),
            (9, 1, "検索条件"),
            (9, 17, "tb2.状態 = 'ACTIVE'"));
    }

    /// <summary>
    /// INSERT SELECTをSELECT表とデータ移送表のハイブリッドへ変換することを確認
    /// </summary>
    [TestMethod]
    public void Build_CreatesInsertSelectHybridFrames()
    {
        const string sql = """
            INSERT INTO user_archive(ユーザーID, 氏名, 状態)
            SELECT
                tb1.ユーザーID
                , tb1.氏名
                , tb1.状態
            FROM
                users AS tb1
            WHERE
                tb1.削除日時 < @archive_before
            """;
        MappingDefinition[] mappings =
        [
            new("user_archive", "ユーザーアーカイブ", "user_id", "ユーザーID"),
            new("tb1", "ユーザー", "user_id", "ユーザーID")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.AreEqual(13, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "サブクエリ[SQ1]"),
            (2, 1, "参照テーブル: ユーザー[tb1]"),
            (3, 1, "取得項目"),
            (3, 7, "取得項目1"),
            (3, 15, ":"),
            (3, 17, "tb1.ユーザーID"),
            (4, 7, "取得項目2"),
            (4, 15, ":"),
            (4, 17, "tb1.氏名"),
            (5, 7, "取得項目3"),
            (5, 15, ":"),
            (5, 17, "tb1.状態"),
            (6, 1, "検索条件"),
            (6, 17, "tb1.削除日時 < @archive_before"),
            (8, 1, "＜データ移送表＞"),
            (9, 1, "参照テーブル: ユーザーアーカイブ、SQ1"),
            (10, 1, "項目"),
            (10, 19, "移送元"),
            (10, 37, "移送方法ほか"),
            (11, 1, "ユーザーID"),
            (11, 19, "SQ1.ユーザーID"),
            (12, 1, "氏名"),
            (12, 19, "SQ1.氏名"),
            (13, 1, "状態"),
            (13, 19, "SQ1.状態"));
    }

    /// <summary>
    /// 複雑なINSERT SELECTの計算と条件をSELECT表へ分離することを確認
    /// </summary>
    [TestMethod]
    public void Build_CreatesComplexInsertSelectHybridFrames()
    {
        const string sql = """
            insert into user_summary(ユーザーID, 表示名, 注文件数, 作成日時, 作成元区分)
            select
                tb1.ユーザーID
                , coalesce(tb1.氏名, tb1.メール) as display_name
                , count(tb2.注文ID) as order_count
                , sysdatetime()
                , 'BATCH'
            from
                users as tb1
                left join orders as tb2
                    on tb1.ユーザーID = tb2.注文ユーザーID
                    and tb2.状態 = 'PAID'
            where
                tb1.状態 = 'ACTIVE'
            group by
                tb1.ユーザーID
                , tb1.氏名
                , tb1.メール
            having
                count(tb2.注文ID) >= @min_order_count
            """;
        MappingDefinition[] mappings =
        [
            new("user_summary", "ユーザー集計", "", ""),
            new("tb1", "ユーザー", "", ""),
            new("tb2", "注文", "", "")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual(24, plan.RowCount);
        Assert.AreEqual("サブクエリ[SQ1]", CellValue(plan, 1, 1));
        Assert.AreEqual("display_name", CellValue(plan, 4, 17));
        Assert.AreEqual("COALESCE(tb1.氏名, tb1.メール)", CellValue(plan, 4, 32));
        Assert.AreEqual("グループ", CellValue(plan, 12, 1));
        Assert.AreEqual("集計条件", CellValue(plan, 15, 1));
        Assert.AreEqual("COUNT(tb2.注文ID) >= @min_order_count", CellValue(plan, 15, 17));
        Assert.AreEqual("＜データ移送表＞", CellValue(plan, 17, 1));
        Assert.AreEqual("参照テーブル: ユーザー集計、SQ1", CellValue(plan, 18, 1));
        Assert.AreEqual("SQ1.表示名", CellValue(plan, 21, 19));
        Assert.AreEqual("SQ1.注文件数", CellValue(plan, 22, 19));
        Assert.AreEqual("SQ1.作成日時", CellValue(plan, 23, 19));
        Assert.IsNull(CellValue(plan, 21, 37));
    }

    /// <summary>
    /// 単一行のINSERT VALUESを値式付きデータ移送表へ変換することを確認
    /// </summary>
    [TestMethod]
    public void Build_CreatesInsertValuesTransferFrame()
    {
        const string sql = """
            INSERT INTO users(ユーザーID, 氏名, 状態, 作成日時)
            VALUES (@user_id, @name, 'ACTIVE', sysdatetime())
            """;
        MappingDefinition[] mappings =
        [
            new("users", "ユーザー", "", "")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual(7, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "＜データ移送表＞"),
            (2, 1, "参照テーブル: ユーザー"),
            (3, 1, "項目"),
            (3, 19, "移送元"),
            (3, 37, "移送方法ほか"),
            (4, 1, "ユーザーID"),
            (4, 37, "@user_id"),
            (5, 1, "氏名"),
            (5, 37, "@name"),
            (6, 1, "状態"),
            (6, 37, "'ACTIVE'"),
            (7, 1, "作成日時"),
            (7, 37, "sysdatetime()"));
    }

    /// <summary>
    /// SELECT INTOをサブクエリ・DB定義・データ移送表のハイブリッドへ変換することを確認
    /// </summary>
    [TestMethod]
    public void Build_CreatesSelectIntoHybridFrames()
    {
        const string sql = """
            select
                tb1.ユーザーID
                , tb1.メール
            into user_export
            from
                users as tb1
            """;
        MappingDefinition[] mappings =
        [
            new("user_export", "ユーザー出力", "", ""),
            new("tb1", "ユーザー", "", "")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual(15, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "サブクエリ[SQ1]"),
            (2, 1, "参照テーブル: ユーザー[tb1]"),
            (3, 1, "取得項目"),
            (3, 7, "取得項目1"),
            (3, 15, ":"),
            (3, 17, "tb1.ユーザーID"),
            (4, 7, "取得項目2"),
            (4, 15, ":"),
            (4, 17, "tb1.メール"),
            (6, 1, "＜DB入出力項目定義＞"),
            (7, 1, "参照テーブル: SQ1"),
            (8, 1, "取得項目"),
            (8, 7, "取得項目1"),
            (8, 15, ":"),
            (8, 17, "SQ1.ユーザーID"),
            (9, 7, "取得項目2"),
            (9, 15, ":"),
            (9, 17, "SQ1.メール"),
            (11, 1, "＜データ移送表＞"),
            (12, 1, "参照テーブル: ユーザー出力、SQ1"),
            (13, 1, "項目"),
            (13, 19, "移送元"),
            (13, 37, "移送方法ほか"),
            (14, 1, "ユーザーID"),
            (14, 19, "SQ1.ユーザーID"),
            (15, 1, "メール"),
            (15, 19, "SQ1.メール"));
    }

    /// <summary>
    /// SELECT INTOの列エイリアスを変換定義の和名へ解決することを確認
    /// </summary>
    [TestMethod]
    public void Build_ResolvesSelectIntoAliasNamesAcrossHybridFrames()
    {
        const string sql = """
            select
                tb1.ユーザーID
                , coalesce(tb1.氏名, tb1.メール) as display_name
                , count(tb1.ユーザーID) as user_count
            into user_summary
            from
                users as tb1
            group by
                tb1.ユーザーID
                , tb1.氏名
                , tb1.メール
            """;
        MappingDefinition[] mappings =
        [
            new("user_summary", "ユーザー集計", "", ""),
            new("tb1", "ユーザー", "", ""),
            new("-", "", "display_name", "表示名"),
            new("-", "", "user_count", "ユーザー件数")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual(2, plan.Cells.Count(cell => cell.Value == "SQ1.表示名"));
        Assert.AreEqual(2, plan.Cells.Count(cell => cell.Value == "SQ1.ユーザー件数"));
        Assert.IsTrue(plan.Cells.Any(cell => cell.Column == 1 && cell.Value == "表示名"));
        Assert.IsTrue(plan.Cells.Any(cell => cell.Column == 1 && cell.Value == "ユーザー件数"));
        Assert.AreEqual("参照テーブル: ユーザー集計、SQ1", plan.Cells.Single(cell => cell.Value.StartsWith("参照テーブル: ユーザー集計", StringComparison.Ordinal)).Value);
    }

    /// <summary>
    /// DELETEを移送行なしの条件表へ変換することを確認
    /// </summary>
    [TestMethod]
    public void Build_CreatesDeleteConditionFrameWithoutTransferRows()
    {
        const string sql = """
            DELETE tb1
            FROM
                users AS tb1
            WHERE
                tb1.削除日時 < @delete_before
                AND tb1.状態 = 'INACTIVE'
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "status", "状態")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.AreEqual(4, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "＜データ移送表＞"),
            (2, 1, "参照テーブル: ユーザー[tb1]"),
            (3, 1, "検索条件"),
            (3, 17, "tb1.削除日時 < @delete_before"),
            (4, 7, "AND"),
            (4, 17, "tb1.状態 = 'INACTIVE'"));
    }

    /// <summary>
    /// EXISTSを含むDELETEをサブクエリ表とデータ移送表へ分離することを確認
    /// </summary>
    [TestMethod]
    public void Build_CreatesDeleteExistsHybridFrames()
    {
        const string sql = """
            DELETE tb1
            FROM
                users AS tb1
            WHERE
                EXISTS (
                    SELECT
                        1
                    FROM
                        orders AS tb2
                    WHERE
                        tb2.注文ユーザーID = tb1.ユーザーID
                )
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "", ""),
            new("tb2", "注文", "", "")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual(8, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "サブクエリ[SQ1]"),
            (2, 1, "参照テーブル: ユーザー[tb1]、注文[tb2]"),
            (3, 1, "取得項目"),
            (3, 7, "取得項目1"),
            (3, 15, ":"),
            (3, 17, "1"),
            (4, 1, "検索条件"),
            (4, 17, "tb2.注文ユーザーID = tb1.ユーザーID"),
            (6, 1, "＜データ移送表＞"),
            (7, 1, "参照テーブル: ユーザー[tb1]、SQ1"),
            (8, 1, "検索条件"),
            (8, 17, "EXISTS (SQ1)"));
    }

    /// <summary>
    /// UPDATE FROMを移送・結合・条件表へ変換することを確認
    /// </summary>
    [TestMethod]
    public void Build_CreatesUpdateFromTransferAndJoinFrame()
    {
        const string sql = """
            UPDATE tb1
            SET
                状態 = @status
                , 更新日時 = sysdatetime()
            FROM
                users AS tb1
                INNER JOIN orders AS tb2
                    ON tb1.ユーザーID = tb2.注文ユーザーID
            WHERE
                tb2.状態 = @order_status
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "user_id", "ユーザーID"),
            new("tb2", "注文", "user_id", "注文ユーザーID")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.AreEqual(8, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "＜データ移送表＞"),
            (2, 1, "参照テーブル: ユーザー[tb1]、注文[tb2]"),
            (3, 1, "項目"),
            (3, 19, "移送元"),
            (3, 37, "移送方法ほか"),
            (4, 1, "状態"),
            (4, 37, "@status"),
            (5, 1, "更新日時"),
            (5, 37, "sysdatetime()"),
            (6, 1, "結合条件"),
            (6, 17, "＜ユーザー[tb1] INNER JOIN 注文[tb2]＞"),
            (7, 17, "tb1.ユーザーID = tb2.注文ユーザーID"),
            (8, 1, "検索条件"),
            (8, 17, "tb2.状態 = @order_status"));
    }

    /// <summary>
    /// テーブル列を参照する更新式は移送元へ出力することを確認
    /// </summary>
    [TestMethod]
    public void Build_WritesColumnBasedUpdateExpressionAsTransferSource()
    {
        const string sql = """
            UPDATE tb1
            SET
                氏名 = tb2.氏名
            FROM
                users AS tb1
                INNER JOIN import_users AS tb2
                    ON tb1.ユーザーID = tb2.ユーザーID
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "name", "氏名"),
            new("tb2", "取込ユーザー", "name", "氏名")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.AreEqual("tb2.氏名", CellValue(plan, 4, 19));
        Assert.IsNull(CellValue(plan, 4, 37));
    }

    /// <summary>
    /// UPDATE SETのスカラーサブクエリをSELECT表とSQ参照へ分離することを確認
    /// </summary>
    [TestMethod]
    public void Build_CreatesUpdateSetSubqueryHybridFrames()
    {
        const string sql = """
            UPDATE tb1
            SET
                最終注文金額 = (
                    SELECT
                        TOP(1) tb2.金額
                    FROM
                        orders AS tb2
                    WHERE
                        tb2.注文ユーザーID = tb1.ユーザーID
                    ORDER BY
                        tb2.注文日時 DESC
                )
            FROM
                users AS tb1
            WHERE
                tb1.状態 = 'ACTIVE'
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "", ""),
            new("tb2", "注文", "", "")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual(12, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "サブクエリ[SQ1]"),
            (2, 1, "参照テーブル: ユーザー[tb1]、注文[tb2]"),
            (3, 1, "取得件数"),
            (3, 7, "1"),
            (4, 1, "取得項目"),
            (4, 7, "取得項目1"),
            (4, 15, ":"),
            (4, 17, "tb2.金額"),
            (5, 1, "検索条件"),
            (5, 17, "tb2.注文ユーザーID = tb1.ユーザーID"),
            (6, 1, "並び順"),
            (6, 7, "ソートキー1"),
            (6, 15, ":"),
            (6, 17, "tb2.注文日時(降順)"),
            (8, 1, "＜データ移送表＞"),
            (9, 1, "参照テーブル: ユーザー[tb1]、SQ1"),
            (10, 1, "項目"),
            (10, 19, "移送元"),
            (10, 37, "移送方法ほか"),
            (11, 1, "最終注文金額"),
            (11, 19, "SQ1.金額"),
            (12, 1, "検索条件"),
            (12, 17, "tb1.状態 = 'ACTIVE'"));
    }

    /// <summary>
    /// 派生SELECT付きUPDATEを派生表とデータ移送表へ分離することを確認
    /// </summary>
    [TestMethod]
    public void Build_CreatesUpdateDerivedSelectHybridFrames()
    {
        const string sql = """
            UPDATE tb1
            SET
                状態 = sq.最新状態
            FROM
                users AS tb1
                INNER JOIN (
                    SELECT
                        tb2.注文ユーザーID
                        , MAX(tb2.状態) AS 最新状態
                    FROM
                        order_status_history AS tb2
                    GROUP BY
                        tb2.注文ユーザーID
                ) AS sq
                    ON tb1.ユーザーID = sq.注文ユーザーID
            WHERE
                tb1.状態 = 'ACTIVE'
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "", ""),
            new("tb2", "注文状態履歴", "", "")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual("サブクエリ[sq]", CellValue(plan, 1, 1));
        Assert.AreEqual("参照テーブル: 注文状態履歴[tb2]", CellValue(plan, 2, 1));
        Assert.AreEqual("＜データ移送表＞", CellValue(plan, 7, 1));
        Assert.AreEqual("参照テーブル: ユーザー[tb1]、sq", CellValue(plan, 8, 1));
        Assert.AreEqual("sq.最新状態", CellValue(plan, 10, 19));
        Assert.AreEqual("＜ユーザー[tb1] INNER JOIN sq＞", CellValue(plan, 11, 17));
        Assert.AreEqual("tb1.ユーザーID = sq.注文ユーザーID", CellValue(plan, 12, 17));
        Assert.AreEqual("tb1.状態 = 'ACTIVE'", CellValue(plan, 13, 17));
    }

    /// <summary>
    /// 一時テーブルの物理名から和名を解決することを確認
    /// </summary>
    [TestMethod]
    public void Build_ResolvesAliasedTemporaryTableByPhysicalName()
    {
        const string sql = """
            SELECT
                t.ID
            FROM
                #users AS t
            """;
        MappingDefinition[] mappings =
        [
            new("#users", "一時ユーザー", "id", "ID")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual("参照テーブル: 一時ユーザー[t]", CellValue(plan, 2, 1));
    }

    /// <summary>
    /// 未対応のINSERT EXECUTEを行単位で出力しフォールバック原因を添えることを確認
    /// </summary>
    [TestMethod]
    public void Build_WritesUnsupportedInsertExecuteByLineWithFallbackReason()
    {
        const string sql = "INSERT INTO users (id)\r\nEXEC dbo.GetUserIds";

        var plan = OutputSheetPlanBuilder.Build(sql, []);

        Assert.IsTrue(plan.IsFallback);
        Assert.AreEqual("未対応のINSERT形式: EXECUTE", plan.FallbackReason);
        Assert.AreEqual(4, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "INSERT INTO users (id)"),
            (2, 1, "EXEC dbo.GetUserIds"),
            (4, 1, "フォールバック原因: 未対応のINSERT形式: EXECUTE（対象クエリ: アウトプットシート 1～2行目）"));
    }

    /// <summary>
    /// 派生テーブルJOINを名前付きサブクエリと全体クエリへ展開することを確認
    /// </summary>
    [TestMethod]
    public void Build_WritesDerivedTableJoinAsNamedSubquery()
    {
        const string sql = """
            SELECT
                tb1.ID
            FROM
                users AS tb1
                INNER JOIN (
                    SELECT
                        user_id
                    FROM
                        orders
                ) AS sq
                    ON tb1.ID = sq.user_id
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "", ""),
            new("orders", "注文", "", "")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual("サブクエリ[sq]", CellValue(plan, 1, 1));
        Assert.AreEqual("参照テーブル: 注文[orders]", CellValue(plan, 2, 1));
        Assert.AreEqual("＜DB入出力項目定義＞", CellValue(plan, 5, 1));
        Assert.AreEqual("参照テーブル: ユーザー[tb1]、sq", CellValue(plan, 6, 1));
        Assert.AreEqual("＜ユーザー[tb1] INNER JOIN sq＞", CellValue(plan, 8, 17));
        Assert.AreEqual("tb1.ID = sq.user_id", CellValue(plan, 9, 17));
    }

    /// <summary>
    /// ネストした相関サブクエリへ外側テーブルと子サブクエリを記載することを確認
    /// </summary>
    [TestMethod]
    public void Build_WritesCorrelatedReferencesForNestedSubqueries()
    {
        const string sql = """
            select
                tb1.ユーザーID
                , tb1.氏名
            from
                users as tb1
            where
                exists (
                    select
                        1
                    from
                        orders as tb2
                    where
                        tb2.注文ユーザーID = tb1.ユーザーID
                        and tb2.金額 > (
                            select
                                avg(tb3.金額)
                            from
                                orders as tb3
                            where
                                tb3.状態 = tb2.状態
                        )
                )
            """;
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "", ""),
            new("tb2", "注文", "", ""),
            new("tb3", "注文", "", "")
        ];

        var plan = OutputSheetPlanBuilder.Build(sql, mappings);

        Assert.IsFalse(plan.IsFallback);
        Assert.AreEqual("参照テーブル: 注文[tb2]、注文[tb3]", CellValue(plan, 2, 1));
        Assert.AreEqual("AVG(tb3.金額)", CellValue(plan, 3, 17));
        Assert.AreEqual("参照テーブル: ユーザー[tb1]、注文[tb2]、SQ1", CellValue(plan, 7, 1));
        Assert.AreEqual("tb2.金額 > (SQ1)", CellValue(plan, 10, 17));
        Assert.AreEqual("EXISTS (SQ2)", CellValue(plan, 16, 17));
    }

    /// <summary>
    /// 指定セルの値を取得
    /// </summary>
    private static string? CellValue(OutputSheetPlan plan, int row, int column)
    {
        return plan.Cells
            .SingleOrDefault(cell => cell.Row == row && cell.Column == column)
            ?.Value;
    }

    /// <summary>
    /// 描画計画の非空セルを順序に依存せず比較
    /// </summary>
    private static void AssertCells(OutputSheetPlan plan, params (int Row, int Column, string Value)[] expected)
    {
        var actual = plan.Cells
            .OrderBy(cell => cell.Row)
            .ThenBy(cell => cell.Column)
            .Select(cell => (cell.Row, cell.Column, cell.Value))
            .ToArray();
        var sortedExpected = expected
            .OrderBy(cell => cell.Row)
            .ThenBy(cell => cell.Column)
            .ToArray();

        CollectionAssert.AreEqual(sortedExpected, actual);
    }
}
