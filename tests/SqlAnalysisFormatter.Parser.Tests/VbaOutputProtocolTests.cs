using SqlAnalysisFormatter.Parser;

namespace SqlAnalysisFormatter.Parser.Tests;

/// <summary>
/// VBA連携用の描画計画・変換定義プロトコルテスト
/// </summary>
[TestClass]
public sealed class VbaOutputProtocolTests
{
    /// <summary>
    /// 描画計画をVBA用プロトコルへ直列化できることを確認
    /// </summary>
    [TestMethod]
    public void SerializePlan_WritesCellsSectionsAndEscapedText()
    {
        var plan = new OutputSheetPlan(
            [
                new OutputCell(1, 1, "見出し"),
                new OutputCell(3, 17, "line1\r\nline2\\value\tend")
            ],
            [
                new OutputSection(OutputSectionKind.Reference, 2, 2),
                new OutputSection(OutputSectionKind.Standard, 3, 4)
            ],
            4,
            false);

        var text = VbaOutputProtocol.SerializePlan(plan);

        var expected = string.Join(
            "\r\n",
            "SAF_OUTPUT_PLAN\t1\t4\t0",
            "C\t1\t1\t見出し",
            "C\t3\t17\tline1\\r\\nline2\\\\value\\tend",
            "S\tREFERENCE\t2\t2",
            "S\tSTANDARD\t3\t4");
        Assert.AreEqual(expected, text);
    }

    /// <summary>
    /// エスケープ済み和名定義を復元できることを確認
    /// </summary>
    [TestMethod]
    public void ParseMappings_RestoresEscapedFields()
    {
        var text = string.Join(
            "\r\n",
            "SAF_MAPPINGS\t1",
            "M\ttb1\tユーザー\tuser_id\tユーザーID",
            "M\ttb\\t2\t注\\\\文\tname\t氏\\n名");

        var mappings = VbaOutputProtocol.ParseMappings(text);

        CollectionAssert.AreEqual(
            new[]
            {
                new MappingDefinition("tb1", "ユーザー", "user_id", "ユーザーID"),
                new MappingDefinition("tb\t2", "注\\文", "name", "氏\n名")
            },
            mappings.ToArray());
    }
}
