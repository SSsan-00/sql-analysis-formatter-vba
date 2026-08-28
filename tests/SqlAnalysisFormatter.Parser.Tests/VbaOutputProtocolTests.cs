using SqlAnalysisFormatter.Parser;

namespace SqlAnalysisFormatter.Parser.Tests;

/// <summary>
/// VBA連携用の描画計画・変換定義プロトコルテスト
/// </summary>
[TestClass]
public sealed class VbaOutputProtocolTests
{
    /// <summary>
    /// フォールバック理由と原因行をVBA通知用の診断行へ直列化することを確認
    /// </summary>
    [TestMethod]
    public void SerializePlan_WritesFallbackDiagnostic()
    {
        var plan = new OutputSheetPlan(
            [new OutputCell(1, 1, "SELECT * FROM users WHERE;")],
            [],
            3,
            true,
            "WHEREの近くに正しくない構文があります。",
            8,
            8,
            FallbackSourceStartLine: 2,
            FallbackSourceEndLine: 2);

        var text = VbaOutputProtocol.SerializePlan(plan);

        StringAssert.StartsWith(text, "SAF_OUTPUT_PLAN\t5\t3\t1");
        StringAssert.Contains(
            text,
            "F\t2\t2\tWHEREの近くに正しくない構文があります。");
    }

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
                new OutputSection(OutputSectionKind.Standard, 3, 4),
                new OutputSection(OutputSectionKind.TransferGroup, 5, 7)
            ],
            4,
            false,
            ReplacementQualifications:
            [
                new OutputReplacementQualification(2, 8, "名前", "tb1.名前")
            ],
            InputTableIds: ["users"],
            OutputTableIds: ["#wkuser"],
            TransformedQueryLines:
            [
                new OutputTransformedQueryLine(3, "WHERE tb1.状態 = 1"),
                new OutputTransformedQueryLine(2, "SELECT tb1.名前\\値\tAS 表示名")
            ],
            ReplacementValues:
            [
                new OutputReplacementValue(3, 4, "tb1.状態"),
                new OutputReplacementValue(2, 8, "tb1.名前\\値\r\n表示名")
            ]);

        var text = VbaOutputProtocol.SerializePlan(plan);

        var expected = string.Join(
            "\r\n",
            "SAF_OUTPUT_PLAN\t5\t4\t0",
            "C\t1\t1\t見出し",
            "C\t3\t17\tline1\\r\\nline2\\\\value\\tend",
            "Q\t2\t8\t名前\ttb1.名前",
            "R\t2\tSELECT tb1.名前\\\\値\\tAS 表示名",
            "R\t3\tWHERE tb1.状態 = 1",
            "V\t2\t8\ttb1.名前\\\\値\\r\\n表示名",
            "V\t3\t4\ttb1.状態",
            "T\tINPUT\tusers",
            "T\tOUTPUT\t#wkuser",
            "S\tREFERENCE\t2\t2",
            "S\tSTANDARD\t3\t4",
            "S\tTRANSFER_GROUP\t5\t7");
        Assert.AreEqual(expected, text);
    }

    /// <summary>
    /// 先頭がフォールバックでも後続の対応ステートメントから得たテーブル情報を直列化することを確認
    /// </summary>
    [TestMethod]
    public void SerializePlan_WritesPartialTableUsageForFallbackPlan()
    {
        const string sql = """
            CREATE INDEX IX_ignored_id ON dbo.ignored_table(id);
            SELECT u.id FROM dbo.users AS u;
            """;
        var plan = OutputSheetPlanBuilder.Build(sql, []);

        var text = VbaOutputProtocol.SerializePlan(plan);

        Assert.IsTrue(plan.IsFallback);
        StringAssert.Contains(text, "T\tINPUT\tusers");
        Assert.IsFalse(text.Contains(
            "T\tINPUT\tignored_table",
            StringComparison.OrdinalIgnoreCase));
        Assert.IsFalse(text.Contains(
            "T\tOUTPUT\tignored_table",
            StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// エスケープ済み和名定義を復元できることを確認
    /// </summary>
    [TestMethod]
    public void ParseMappings_RestoresEscapedFields()
    {
        var text = string.Join(
            "\r\n",
            "SAF_MAPPINGS\t2",
            "M\ttb1\tユーザー\tuser_id\tユーザーID\t__SAF_FIELD_R000002__",
            "M\ttb\\t2\t注\\\\文\tname\t氏\\n名\t__SAF_FIELD_R000003__");

        var mappings = VbaOutputProtocol.ParseMappings(text);

        CollectionAssert.AreEqual(
            new[]
            {
                new MappingDefinition("tb1", "ユーザー", "user_id", "ユーザーID", "__SAF_FIELD_R000002__"),
                new MappingDefinition("tb\t2", "注\\文", "name", "氏\n名", "__SAF_FIELD_R000003__")
            },
            mappings.ToArray());
    }

    /// <summary>
    /// 旧形式の変換定義も引き続き読み込めることを確認
    /// </summary>
    [TestMethod]
    public void ParseMappings_AcceptsLegacyVersionOne()
    {
        const string text = "SAF_MAPPINGS\t1\r\nM\ttb1\tユーザー\tuser_id\tユーザーID";

        var mappings = VbaOutputProtocol.ParseMappings(text);

        CollectionAssert.AreEqual(
            new[] { new MappingDefinition("tb1", "ユーザー", "user_id", "ユーザーID") },
            mappings.ToArray());
    }
}
