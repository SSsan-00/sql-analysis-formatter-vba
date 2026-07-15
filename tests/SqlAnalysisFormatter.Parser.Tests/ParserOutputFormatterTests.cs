using SqlAnalysisFormatter.Parser;

namespace SqlAnalysisFormatter.Parser.Tests;

/// <summary>
/// VBA連携用出力形式のテスト
/// </summary>
[TestClass]
public sealed class ParserOutputFormatterTests
{
    /// <summary>
    /// 複数ブロックをレコード区切りで連結することを確認
    /// </summary>
    [TestMethod]
    public void ToVbaBlocks_JoinsBlocksWithRecordSeparator()
    {
        var result = new ParseResult(
            true,
            "SELECT",
            [
                new QueryBlock("SUBQUERY", "select 1", 0, 8),
                new QueryBlock("WHOLE", "select * from users", 0, 19)
            ],
            []);

        var formatted = ParserOutputFormatter.ToVbaBlocks(result);

        Assert.AreEqual($"select 1{ParserOutputFormatter.VbaBlockSeparator}select * from users", formatted);
    }

    /// <summary>
    /// 解析失敗時もフォールバックブロックを維持することを確認
    /// </summary>
    [TestMethod]
    public void ToVbaBlocks_KeepsFallbackBlockWhenParseFails()
    {
        var result = TsqlAstParser.Parse("select from");

        var formatted = ParserOutputFormatter.ToVbaBlocks(result);

        Assert.AreEqual("select from", formatted);
    }
}
