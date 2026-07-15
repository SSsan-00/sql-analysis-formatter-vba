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
    /// INSERT SELECTをデータ移送表へ変換することを確認
    /// </summary>
    [TestMethod]
    public void Build_CreatesInsertSelectTransferFrame()
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

        Assert.AreEqual(7, plan.RowCount);
        AssertCells(
            plan,
            (1, 1, "＜データ移送表＞"),
            (2, 1, "参照テーブル: ユーザーアーカイブ、ユーザー[tb1]"),
            (3, 1, "項目"),
            (3, 19, "移送元"),
            (3, 37, "移送方法ほか"),
            (4, 1, "ユーザーID"),
            (4, 19, "tb1.ユーザーID"),
            (5, 1, "氏名"),
            (5, 19, "tb1.氏名"),
            (6, 1, "状態"),
            (6, 19, "tb1.状態"),
            (7, 1, "検索条件"),
            (7, 17, "tb1.削除日時 < @archive_before"));
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
            (4, 19, "@status"),
            (5, 1, "更新日時"),
            (5, 19, "sysdatetime()"),
            (6, 1, "結合条件"),
            (6, 17, "＜ユーザー[tb1] INNER JOIN 注文[tb2]＞"),
            (7, 17, "tb1.ユーザーID = tb2.注文ユーザーID"),
            (8, 1, "検索条件"),
            (8, 17, "tb2.状態 = @order_status"));
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
