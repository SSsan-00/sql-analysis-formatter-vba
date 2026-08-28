using SqlAnalysisFormatter.Parser;

namespace SqlAnalysisFormatter.Parser.Tests;

/// <summary>
/// parser専用フィールドIDの最終表示値への復元テスト
/// </summary>
[TestClass]
public sealed class ParserFieldIdentifierRestorerTests
{
    /// <summary>
    /// 描画セル、診断、SQL解析シート向けの全出力値を同じ和名へ復元
    /// </summary>
    [TestMethod]
    public void Restore_RestoresParserFieldIdsAcrossAllPlanOutputs()
    {
        const string parserFieldId = "__SAF_FIELD_R000002__";
        var plan = new OutputSheetPlan(
            [new OutputCell(1, 1, $"tb1.[{parserFieldId}]")],
            [],
            1,
            true,
            $"未対応: {parserFieldId}",
            ReplacementQualifications:
            [
                new OutputReplacementQualification(
                    1,
                    3,
                    parserFieldId,
                    $"tb1.[{parserFieldId}]")
            ],
            TransformedQueryLines:
            [new OutputTransformedQueryLine(1, $"SELECT tb1.[{parserFieldId}]")],
            ReplacementValues:
            [new OutputReplacementValue(1, 3, $"tb1.{parserFieldId}")]);
        MappingDefinition[] mappings =
        [
            new("tb1", "ユーザー", "name", "表示名/名称", parserFieldId)
        ];

        var restored = ParserFieldIdentifierRestorer.Restore(plan, mappings);

        Assert.AreEqual("tb1.表示名/名称", restored.Cells[0].Value);
        Assert.AreEqual("未対応: 表示名/名称", restored.FallbackReason);
        Assert.AreEqual(
            "表示名/名称",
            restored.ReplacementQualifications![0].OriginalValue);
        Assert.AreEqual(
            "tb1.表示名/名称",
            restored.ReplacementQualifications[0].QualifiedValue);
        Assert.AreEqual(
            "SELECT tb1.表示名/名称",
            restored.TransformedQueryLines[0].Value);
        Assert.AreEqual(
            "tb1.表示名/名称",
            restored.ReplacementValues[0].Value);
    }

    /// <summary>
    /// 復元定義がない場合は既定の空コレクションを持つ元計画を維持
    /// </summary>
    [TestMethod]
    public void Restore_ReturnsOriginalPlanWhenNoParserFieldMappingExists()
    {
        var plan = new OutputSheetPlan([], [], 0, false);

        var restored = ParserFieldIdentifierRestorer.Restore(plan, []);

        Assert.AreSame(plan, restored);
        Assert.IsEmpty(restored.TransformedQueryLines);
        Assert.IsEmpty(restored.ReplacementValues);
    }
}
