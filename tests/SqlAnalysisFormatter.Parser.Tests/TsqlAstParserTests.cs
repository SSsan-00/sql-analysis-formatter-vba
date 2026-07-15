using SqlAnalysisFormatter.Parser;

namespace SqlAnalysisFormatter.Parser.Tests;

/// <summary>
/// ScriptDom AST解析の基本仕様テスト
/// </summary>
[TestClass]
public sealed class TsqlAstParserTests
{
    /// <summary>
    /// 先頭文のCRUD種別を分類できることを確認
    /// </summary>
    [TestMethod]
    [DataRow("select users.user_id from users", "SELECT")]
    [DataRow("insert into users(user_id) values (1)", "INSERT")]
    [DataRow("update users set name = 'x'", "UPDATE")]
    [DataRow("delete from users where user_id = 1", "DELETE")]
    public void Parse_ClassifiesTopLevelCrud(string sql, string expectedQueryType)
    {
        var result = TsqlAstParser.Parse(sql);

        Assert.IsTrue(result.Ok, string.Join(Environment.NewLine, result.ParseErrors));
        Assert.AreEqual(expectedQueryType, result.QueryType);
        Assert.HasCount(1, result.Blocks);
        Assert.AreEqual(sql, result.Blocks[0].Text);
    }

    /// <summary>
    /// ネストしたサブクエリを内側から返すことを確認
    /// </summary>
    [TestMethod]
    public void Parse_ReturnsNestedSubqueriesInsideOut()
    {
        const string sql = """
            with high_value_orders as (
                select
                    orders.user_id
                    , orders.amount
                from
                    orders
                where
                    orders.amount > 1000
                    and exists (
                        select
                            1
                        from
                            order_items
                        where
                            order_items.order_id = orders.order_id
                    )
            )
            select
                users.user_id
                , high_value_orders.amount
            from
                users
                inner join high_value_orders
                    on high_value_orders.user_id = users.user_id
            """;

        var result = TsqlAstParser.Parse(sql);

        Assert.IsTrue(result.Ok, string.Join(Environment.NewLine, result.ParseErrors));
        Assert.AreEqual("SELECT", result.QueryType);
        Assert.HasCount(3, result.Blocks);
        StringAssert.StartsWith(result.Blocks[0].Text, "select");
        StringAssert.Contains(result.Blocks[0].Text, "order_items");
        StringAssert.StartsWith(result.Blocks[1].Text, "select");
        StringAssert.Contains(result.Blocks[1].Text, "exists");
        StringAssert.StartsWith(result.Blocks[2].Text, "with high_value_orders");
    }

    /// <summary>
    /// 未対応文をSQL全体の単一ブロックとして返すことを確認
    /// </summary>
    [TestMethod]
    public void Parse_ReturnsUnsupportedQueryAsSingleBlock()
    {
        const string sql = """
            exec dbo.refresh_user_summary
                @target_date = '2026-07-12'
            """;

        var result = TsqlAstParser.Parse(sql);

        Assert.IsTrue(result.Ok, string.Join(Environment.NewLine, result.ParseErrors));
        Assert.AreEqual("EXEC", result.QueryType);
        Assert.HasCount(1, result.Blocks);
        Assert.AreEqual(sql.Trim(), result.Blocks[0].Text);
    }

    /// <summary>
    /// 構文エラー時に元SQLへフォールバックすることを確認
    /// </summary>
    [TestMethod]
    public void Parse_FallsBackToOriginalSqlWhenParseFails()
    {
        const string sql = "select from";

        var result = TsqlAstParser.Parse(sql);

        Assert.IsFalse(result.Ok);
        Assert.AreEqual("UNKNOWN", result.QueryType);
        Assert.HasCount(1, result.Blocks);
        Assert.AreEqual(sql, result.Blocks[0].Text);
        Assert.IsGreaterThan(0, result.ParseErrors.Count);
    }
}
