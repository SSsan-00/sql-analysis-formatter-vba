namespace SqlAnalysisFormatter.Parser;

/// <summary>
/// アウトプットシートへ設定するセル
/// </summary>
public sealed record OutputCell(int Row, int Column, string Value);

/// <summary>
/// VBA側で共通書式を適用するセクション種別
/// </summary>
public enum OutputSectionKind
{
    Reference,
    Standard,
    Transfer,
    TransferGroup,
    Separator
}

/// <summary>
/// 共通書式を適用する行範囲
/// </summary>
public sealed record OutputSection(OutputSectionKind Kind, int StartRow, int EndRow);

/// <summary>
/// SQL解析シートの変換内容へ反映するプレフィックス補完
/// </summary>
public sealed record OutputReplacementQualification(
    int QueryLine,
    int Order,
    string OriginalValue,
    string QualifiedValue);

/// <summary>
/// アウトプットシート全体の描画計画
/// </summary>
public sealed record OutputSheetPlan(
    IReadOnlyList<OutputCell> Cells,
    IReadOnlyList<OutputSection> Sections,
    int RowCount,
    bool IsFallback,
    string? FallbackReason = null,
    int? FallbackQueryStartRow = null,
    int? FallbackQueryEndRow = null,
    IReadOnlyList<OutputReplacementQualification>? ReplacementQualifications = null,
    IReadOnlyList<string>? InputTableIds = null,
    IReadOnlyList<string>? OutputTableIds = null)
{
    public IReadOnlyList<string> InputTableIds { get; init; } = InputTableIds ?? [];

    public IReadOnlyList<string> OutputTableIds { get; init; } = OutputTableIds ?? [];
}
