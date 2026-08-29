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
/// SQL解析シートのB列へ返す変換後クエリの論理行
/// </summary>
public sealed record OutputTransformedQueryLine(int QueryLine, string Value);

/// <summary>
/// SQL解析シートのC列以降へ返す最終変換値
/// </summary>
public sealed record OutputReplacementValue(
    int QueryLine,
    int Order,
    string Value);

/// <summary>
/// テーブル一覧の名称で安全に補完できる帳票上の物理テーブル表示
/// </summary>
public sealed record OutputTableNameReference(
    int Row,
    int Column,
    string SourceValue,
    string PhysicalTableId,
    string ReplacementSuffix);

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
    IReadOnlyList<string>? OutputTableIds = null,
    int? FallbackSourceStartLine = null,
    int? FallbackSourceEndLine = null,
    IReadOnlyList<OutputTransformedQueryLine>? TransformedQueryLines = null,
    IReadOnlyList<OutputReplacementValue>? ReplacementValues = null,
    IReadOnlyList<OutputTableNameReference>? TableNameReferences = null)
{
    public IReadOnlyList<string> InputTableIds { get; init; } = InputTableIds ?? [];

    public IReadOnlyList<string> OutputTableIds { get; init; } = OutputTableIds ?? [];

    public IReadOnlyList<OutputTransformedQueryLine> TransformedQueryLines { get; init; } =
        TransformedQueryLines ?? [];

    public IReadOnlyList<OutputReplacementValue> ReplacementValues { get; init; } =
        ReplacementValues ?? [];

    public IReadOnlyList<OutputTableNameReference> TableNameReferences { get; init; } =
        TableNameReferences ?? [];
}
