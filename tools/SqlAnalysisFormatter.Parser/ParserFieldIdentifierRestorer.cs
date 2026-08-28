using System.Text;

namespace SqlAnalysisFormatter.Parser;

/// <summary>
/// parser専用フィールドIDを帳票表示用の和名へ復元
/// </summary>
internal static class ParserFieldIdentifierRestorer
{
    private const string ParserFieldPrefix = "__SAF_FIELD_R";
    private const string ParserFieldSuffix = "__";

    /// <summary>
    /// 描画計画に残ったparser専用IDを和名へ置換
    /// </summary>
    public static OutputSheetPlan Restore(
        OutputSheetPlan plan,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var replacements = mappings
            .Where(mapping =>
                !string.IsNullOrEmpty(mapping.ParserFieldId) &&
                !string.IsNullOrEmpty(mapping.FieldName))
            .GroupBy(mapping => mapping.ParserFieldId, StringComparer.Ordinal)
            .ToDictionary(
                group => group.Key,
                group => group.Last().FieldName,
                StringComparer.Ordinal);
        if (replacements.Count == 0)
        {
            return plan;
        }

        var cells = plan.Cells
            .Select(cell => cell with { Value = RestoreText(cell.Value, replacements) })
            .ToArray();
        var fallbackReason = plan.FallbackReason is null
            ? null
            : RestoreText(plan.FallbackReason, replacements);
        var replacementQualifications = plan.ReplacementQualifications?
            .Select(qualification => qualification with
            {
                OriginalValue = RestoreText(qualification.OriginalValue, replacements),
                QualifiedValue = RestoreText(qualification.QualifiedValue, replacements)
            })
            .ToArray();
        var transformedQueryLines = plan.TransformedQueryLines
            .Select(line => line with { Value = RestoreText(line.Value, replacements) })
            .ToArray();
        var replacementValues = plan.ReplacementValues
            .Select(item => item with { Value = RestoreText(item.Value, replacements) })
            .ToArray();
        return plan with
        {
            Cells = cells,
            FallbackReason = fallbackReason,
            ReplacementQualifications = replacementQualifications,
            TransformedQueryLines = transformedQueryLines,
            ReplacementValues = replacementValues
        };
    }

    /// <summary>
    /// 文字列内の角括弧付き・通常表記のparser専用IDを復元
    /// </summary>
    private static string RestoreText(
        string value,
        IReadOnlyDictionary<string, string> replacements)
    {
        if (!value.Contains(ParserFieldPrefix, StringComparison.Ordinal))
        {
            return value;
        }

        var result = new StringBuilder(value.Length);
        var currentIndex = 0;
        while (currentIndex < value.Length)
        {
            var identifierStart = value.IndexOf(
                ParserFieldPrefix,
                currentIndex,
                StringComparison.Ordinal);
            if (identifierStart < 0)
            {
                result.Append(value, currentIndex, value.Length - currentIndex);
                break;
            }

            var suffixStart = value.IndexOf(
                ParserFieldSuffix,
                identifierStart + ParserFieldPrefix.Length,
                StringComparison.Ordinal);
            if (suffixStart < 0)
            {
                result.Append(value, currentIndex, value.Length - currentIndex);
                break;
            }

            var identifierEnd = suffixStart + ParserFieldSuffix.Length;
            var parserFieldId = value[identifierStart..identifierEnd];
            if (!replacements.TryGetValue(parserFieldId, out var fieldName))
            {
                result.Append(value, currentIndex, identifierEnd - currentIndex);
                currentIndex = identifierEnd;
                continue;
            }

            var tokenStart = identifierStart;
            var tokenEnd = identifierEnd;
            if (identifierStart > currentIndex &&
                value[identifierStart - 1] == '[' &&
                identifierEnd < value.Length &&
                value[identifierEnd] == ']')
            {
                tokenStart--;
                tokenEnd++;
            }

            result.Append(value, currentIndex, tokenStart - currentIndex);
            result.Append(fieldName);
            currentIndex = tokenEnd;
        }

        return result.ToString();
    }
}
