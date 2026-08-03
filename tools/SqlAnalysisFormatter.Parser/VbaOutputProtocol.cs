using System.Text;

namespace SqlAnalysisFormatter.Parser;

/// <summary>
/// VBAとparser exeの間で描画計画と変換定義を受け渡す
/// </summary>
public static class VbaOutputProtocol
{
    private const string PlanHeader = "SAF_OUTPUT_PLAN\t3";
    private const string MappingHeaderV1 = "SAF_MAPPINGS\t1";
    private const string MappingHeaderV2 = "SAF_MAPPINGS\t2";

    /// <summary>
    /// 描画計画をタブ区切り行へ変換
    /// </summary>
    public static string SerializePlan(OutputSheetPlan plan)
    {
        ArgumentNullException.ThrowIfNull(plan);

        var lines = new List<string>
        {
            $"{PlanHeader}\t{plan.RowCount}\t{(plan.IsFallback ? 1 : 0)}"
        };
        lines.AddRange(plan.Cells
            .OrderBy(cell => cell.Row)
            .ThenBy(cell => cell.Column)
            .Select(cell => $"C\t{cell.Row}\t{cell.Column}\t{Escape(cell.Value)}"));
        lines.AddRange((plan.ReplacementQualifications ?? [])
            .OrderBy(item => item.QueryLine)
            .ThenBy(item => item.Order)
            .Select(item =>
                $"Q\t{item.QueryLine}\t{item.Order}\t{Escape(item.OriginalValue)}\t{Escape(item.QualifiedValue)}"));
        lines.AddRange(plan.InputTableIds.Select(tableId => $"T\tINPUT\t{Escape(tableId)}"));
        lines.AddRange(plan.OutputTableIds.Select(tableId => $"T\tOUTPUT\t{Escape(tableId)}"));
        lines.AddRange(plan.Sections.Select(section =>
            $"S\t{SectionKindText(section.Kind)}\t{section.StartRow}\t{section.EndRow}"));
        return string.Join("\r\n", lines);
    }

    /// <summary>
    /// VBA側で読みやすいセクション名へ変換
    /// </summary>
    private static string SectionKindText(OutputSectionKind kind)
    {
        return kind switch
        {
            OutputSectionKind.TransferGroup => "TRANSFER_GROUP",
            _ => kind.ToString().ToUpperInvariant()
        };
    }

    /// <summary>
    /// VBAが出力した変換定義行を復元
    /// </summary>
    public static IReadOnlyList<MappingDefinition> ParseMappings(string text)
    {
        ArgumentNullException.ThrowIfNull(text);

        var lines = text
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n', StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length == 0 ||
            (lines[0] != MappingHeaderV1 && lines[0] != MappingHeaderV2))
        {
            throw new InvalidDataException("変換定義ファイルのヘッダーが不正です。");
        }

        var hasParserFieldId = lines[0] == MappingHeaderV2;
        var expectedFieldCount = hasParserFieldId ? 6 : 5;
        var mappings = new List<MappingDefinition>();
        for (var index = 1; index < lines.Length; index++)
        {
            var fields = lines[index].Split('\t');
            if (fields.Length != expectedFieldCount || fields[0] != "M")
            {
                throw new InvalidDataException($"変換定義ファイルの{index + 1}行目が不正です。");
            }

            mappings.Add(new MappingDefinition(
                Unescape(fields[1]),
                Unescape(fields[2]),
                Unescape(fields[3]),
                Unescape(fields[4]),
                hasParserFieldId ? Unescape(fields[5]) : string.Empty));
        }

        return mappings;
    }

    /// <summary>
    /// 制御文字とバックスラッシュを可逆変換
    /// </summary>
    private static string Escape(string value)
    {
        return value
            .Replace("\\", "\\\\", StringComparison.Ordinal)
            .Replace("\r", "\\r", StringComparison.Ordinal)
            .Replace("\n", "\\n", StringComparison.Ordinal)
            .Replace("\t", "\\t", StringComparison.Ordinal);
    }

    /// <summary>
    /// エスケープ済み文字列を元へ戻す
    /// </summary>
    private static string Unescape(string value)
    {
        var result = new StringBuilder(value.Length);
        for (var index = 0; index < value.Length; index++)
        {
            if (value[index] != '\\' || index + 1 >= value.Length)
            {
                result.Append(value[index]);
                continue;
            }

            index++;
            result.Append(value[index] switch
            {
                'r' => '\r',
                'n' => '\n',
                't' => '\t',
                '\\' => '\\',
                _ => "\\" + value[index]
            });
        }

        return result.ToString();
    }
}
