using Microsoft.SqlServer.TransactSql.ScriptDom;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace SqlAnalysisFormatter.Parser;

/// <summary>
/// T-SQL ASTをアウトプットシートの描画計画へ変換
/// </summary>
public static class OutputSheetPlanBuilder
{
    private const string MissingName = "(和名未取得)";
    private const int WrappedCaseBranchIndentColumns = 8;
    private const int WrappedMultipleCaseLabelIndentColumns = 14;
    private const int MultipleCaseBranchIndentColumns = 6;
    private const int OffsetCaseMarkerColumn = 27;
    private const int OffsetCaseDetailColumn = 28;
    private static readonly AsyncLocal<DisplayAliasContext?> CurrentDisplayAliases = new();
    private static readonly AsyncLocal<List<MissingTableDisplayCandidate>?>
        CurrentMissingTableDisplays = new();

    /// <summary>
    /// 和名変換済みSQLから描画計画を作成
    /// </summary>
    public static OutputSheetPlan Build(string sql, IReadOnlyList<MappingDefinition> mappings)
    {
        ArgumentNullException.ThrowIfNull(sql);
        ArgumentNullException.ThrowIfNull(mappings);

        var previousCandidates = CurrentMissingTableDisplays.Value;
        var candidates = new List<MissingTableDisplayCandidate>();
        CurrentMissingTableDisplays.Value = candidates;
        try
        {
            var plan = BuildCore(sql, mappings);
            return plan with
            {
                TableNameReferences = BuildTableNameReferences(plan, candidates)
            };
        }
        finally
        {
            CurrentMissingTableDisplays.Value = previousCandidates;
        }
    }

    /// <summary>
    /// SQL解析本体を実行し、物理テーブル表示の収集スコープは呼出元で管理
    /// </summary>
    private static OutputSheetPlan BuildCore(
        string sql,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var sourceScript = ParseScript(sql, out var errors);
        if (errors.Count > 0 || sourceScript is null)
        {
            var fallback = AttachTransformedQueryResults(
                CreateParseFallback(sql, errors),
                sql,
                null,
                mappings,
                []);
            return ParserFieldIdentifierRestorer.Restore(fallback, mappings);
        }

        var qualificationResult = QualifyUnqualifiedSelectColumns(sql, sourceScript, mappings);
        var qualifiedSql = qualificationResult.Sql;
        var script = sourceScript;
        if (!string.Equals(qualifiedSql, sql, StringComparison.Ordinal))
        {
            script = ParseScript(qualifiedSql, out errors);
            if (errors.Count > 0 || script is null)
            {
                var fallback = AttachTransformedQueryResults(
                    CreateParseFallback(qualifiedSql, errors),
                    sql,
                    sourceScript,
                    mappings,
                    []);
                return ParserFieldIdentifierRestorer.Restore(fallback, mappings);
            }
        }

        var statements = Statements(script);
        var statementPlans = statements
            .Select(statement => (
                Statement: statement,
                Plan: BuildStatement(qualifiedSql, statement, mappings)))
            .ToArray();
        var plan = statementPlans.Length > 0
            ? statementPlans[0].Plan
            : BuildStatement(qualifiedSql, null, mappings);
        var usage = PhysicalTableUsageCollector.Collect(
            statementPlans
                .Where(result => !result.Plan.IsFallback)
                .Select(result => result.Statement));

        var supportedQualifiedStatements = statementPlans
            .Where(result => !result.Plan.IsFallback)
            .Select(result => result.Statement)
            .ToArray();
        var sourceStatements = Statements(sourceScript);
        var supportedSourceStatements = sourceStatements.Count == statementPlans.Length
            ? sourceStatements
                .Where((_, index) => !statementPlans[index].Plan.IsFallback)
                .ToArray()
            : [];
        var sourceAliasReplacements =
            BinaryDisplayAliasReplacementCollector.Collect(supportedSourceStatements);
        var qualifiedAliasReplacements =
            BinaryDisplayAliasReplacementCollector.Collect(supportedQualifiedStatements);
        var transformedSql = ApplyTextReplacements(
            sql,
            0,
            sourceAliasReplacements);
        plan = plan with
        {
            ReplacementQualifications = qualificationResult.Replacements,
            InputTableIds = usage.InputTableIds,
            OutputTableIds = usage.OutputTableIds,
            TransformedQueryLines = BuildTransformedQueryLines(transformedSql),
            ReplacementValues = BuildReplacementValues(
                qualifiedSql,
                script,
                mappings,
                qualifiedAliasReplacements)
        };
        return ParserFieldIdentifierRestorer.Restore(plan, mappings);
    }

    /// <summary>
    /// 参照テーブル行とJOIN見出しにある一意な物理テーブル表示を補完対象として返す
    /// </summary>
    private static IReadOnlyList<OutputTableNameReference> BuildTableNameReferences(
        OutputSheetPlan plan,
        IReadOnlyList<MissingTableDisplayCandidate> candidates)
    {
        if (candidates.Count == 0)
        {
            return [];
        }

        var uniqueCandidates = candidates
            .GroupBy(candidate => candidate.SourceValue, StringComparer.OrdinalIgnoreCase)
            .Where(group => group
                .Select(candidate => candidate.PhysicalTableId)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(2)
                .Count() == 1)
            .Where(group => group
                .Select(candidate => candidate.ReplacementSuffix)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(2)
                .Count() == 1)
            .Select(group => group.First())
            .OrderByDescending(candidate => candidate.SourceValue.Length)
            .ThenBy(candidate => candidate.SourceValue, StringComparer.Ordinal)
            .ToArray();
        var references = new List<OutputTableNameReference>();
        foreach (var cell in plan.Cells.Where(IsTableNameReferenceCell))
        {
            foreach (var candidate in uniqueCandidates)
            {
                if (!ContainsTableDisplay(cell, candidate))
                {
                    continue;
                }

                references.Add(new OutputTableNameReference(
                    cell.Row,
                    cell.Column,
                    candidate.SourceValue,
                    candidate.PhysicalTableId,
                    candidate.ReplacementSuffix));
            }
        }

        return references;
    }

    /// <summary>
    /// テーブル表示だけを持つ帳票セルを条件式やSQL文字列のセルから分離
    /// </summary>
    private static bool IsTableNameReferenceCell(OutputCell cell)
    {
        if (cell.Column == 1 &&
            cell.Value.StartsWith("参照テーブル: ", StringComparison.Ordinal))
        {
            return true;
        }

        return cell.Column == 17 &&
            cell.Value.StartsWith('＜') &&
            cell.Value.EndsWith('＞') &&
            cell.Value.Contains(" JOIN ", StringComparison.Ordinal);
    }

    /// <summary>
    /// 参照一覧またはJOIN左右の表示要素と完全一致する候補だけを採用
    /// </summary>
    private static bool ContainsTableDisplay(
        OutputCell cell,
        MissingTableDisplayCandidate candidate)
    {
        IEnumerable<string> displays;
        if (cell.Column == 1)
        {
            const string prefix = "参照テーブル: ";
            displays = cell.Value[prefix.Length..].Split('、');
        }
        else
        {
            var joinText = cell.Value[1..^1];
            var delimiter = new[]
                {
                    " INNER JOIN ",
                    " LEFT JOIN ",
                    " RIGHT JOIN ",
                    " FULL JOIN ",
                    " JOIN "
                }
                .First(item => joinText.Contains(item, StringComparison.Ordinal));
            displays = joinText
                .Split(delimiter, 2, StringSplitOptions.None)
                .SelectMany(side => side.Split('、'));
        }

        return displays.Any(display => string.Equals(
            display,
            candidate.SourceValue,
            StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// parserが物理テーブルを一意に把握した未取得表示を構造化して収集
    /// </summary>
    private static string RegisterMissingTableDisplay(
        string sourceValue,
        string physicalTableId,
        string replacementSuffix)
    {
        CurrentMissingTableDisplays.Value?.Add(new MissingTableDisplayCandidate(
            sourceValue,
            physicalTableId,
            replacementSuffix));
        return sourceValue;
    }

    /// <summary>
    /// スクリプト内のステートメントをバッチ順に列挙
    /// </summary>
    private static IReadOnlyList<TSqlStatement> Statements(TSqlScript script)
    {
        return script.Batches
            .SelectMany(batch => batch.Statements)
            .ToArray();
    }

    /// <summary>
    /// フォールバック時にも従来の和名変換結果をB列とC列へ返却
    /// </summary>
    private static OutputSheetPlan AttachTransformedQueryResults(
        OutputSheetPlan plan,
        string sql,
        TSqlScript? script,
        IReadOnlyList<MappingDefinition> mappings,
        IReadOnlyList<SqlTextReplacement> aliasReplacements)
    {
        return plan with
        {
            TransformedQueryLines = BuildTransformedQueryLines(sql),
            ReplacementValues = BuildReplacementValues(
                sql,
                script,
                mappings,
                aliasReplacements)
        };
    }

    /// <summary>
    /// SQLの行数と空行を維持してB列返却用の論理行へ分割
    /// </summary>
    private static IReadOnlyList<OutputTransformedQueryLine> BuildTransformedQueryLines(
        string sql)
    {
        return sql
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Split('\n')
            .Select((value, index) => new OutputTransformedQueryLine(index + 1, value))
            .ToArray();
    }

    /// <summary>
    /// parser専用フィールドIDの各出現を、C列へ返す最終表示値へ変換
    /// </summary>
    private static IReadOnlyList<OutputReplacementValue> BuildReplacementValues(
        string sql,
        TSqlScript? script,
        IReadOnlyList<MappingDefinition> mappings,
        IReadOnlyList<SqlTextReplacement> aliasReplacements)
    {
        var mappingsByParserId = mappings
            .Where(mapping => mapping.ParserFieldId.Length > 0)
            .GroupBy(mapping => mapping.ParserFieldId, StringComparer.Ordinal)
            .ToDictionary(
                group => group.Key,
                group => group.Last(),
                StringComparer.Ordinal);
        if (mappingsByParserId.Count == 0)
        {
            return [];
        }

        var occurrences = mappingsByParserId
            .SelectMany(item => FindTextOccurrences(sql, item.Key)
                .Select(offset => new ParserFieldOccurrence(offset, item.Value)))
            .OrderBy(item => item.Offset)
            .ToArray();
        if (occurrences.Length == 0)
        {
            return [];
        }

        var parserColumns = script is null
            ? []
            : ParserFieldColumnCollector.Collect(script, mappingsByParserId.Keys);
        var replacementsByOffset = aliasReplacements
            .GroupBy(item => item.Offset)
            .ToDictionary(group => group.Key, group => group.Last());
        var positions = SqlTextPositions(sql, occurrences.Select(item => item.Offset));
        var values = new List<OutputReplacementValue>(occurrences.Length);
        for (var index = 0; index < occurrences.Length; index++)
        {
            var occurrence = occurrences[index];
            var column = parserColumns.FirstOrDefault(item =>
                string.Equals(
                    item.FieldIdentifier.Value,
                    occurrence.Mapping.ParserFieldId,
                    StringComparison.Ordinal) &&
                occurrence.Offset >= item.FieldIdentifier.StartOffset &&
                occurrence.Offset < item.FieldIdentifier.StartOffset +
                    item.FieldIdentifier.FragmentLength);
            var value = ReplacementValueFor(
                sql,
                occurrence.Mapping,
                column,
                replacementsByOffset);
            values.Add(new OutputReplacementValue(
                positions[index].Line,
                positions[index].Column,
                value));
        }

        return values;
    }

    /// <summary>
    /// 1件のparser専用IDを元定義、未修飾列補完、表示用別名から最終値へ変換
    /// </summary>
    private static string ReplacementValueFor(
        string sql,
        MappingDefinition mapping,
        ParserFieldColumn? parserColumn,
        IReadOnlyDictionary<int, SqlTextReplacement> aliasReplacements)
    {
        var value = mapping.TableId == "-"
            ? mapping.ParserFieldId
            : mapping.TableId + "." + mapping.ParserFieldId;
        var identifiers = parserColumn?.Column.MultiPartIdentifier?.Identifiers;
        if (identifiers is null || identifiers.Count < 2)
        {
            return value;
        }

        var qualifier = identifiers[^2];
        if (aliasReplacements.TryGetValue(qualifier.StartOffset, out var aliasReplacement))
        {
            return aliasReplacement.Value + "." + mapping.ParserFieldId;
        }

        if (mapping.TableId == "-")
        {
            return RawIdentifierText(sql, qualifier) + "." + mapping.ParserFieldId;
        }

        return value;
    }

    /// <summary>
    /// 識別子の引用符を含む元表記を取得
    /// </summary>
    private static string RawIdentifierText(string sql, Identifier identifier)
    {
        return identifier.StartOffset >= 0 &&
            identifier.FragmentLength > 0 &&
            identifier.StartOffset + identifier.FragmentLength <= sql.Length
                ? sql.Substring(identifier.StartOffset, identifier.FragmentLength)
                : identifier.Value;
    }

    /// <summary>
    /// SQL内にある指定文字列の全開始位置を列挙
    /// </summary>
    private static IEnumerable<int> FindTextOccurrences(string sql, string value)
    {
        var offset = 0;
        while (offset <= sql.Length - value.Length)
        {
            offset = sql.IndexOf(value, offset, StringComparison.Ordinal);
            if (offset < 0)
            {
                yield break;
            }

            yield return offset;
            offset += value.Length;
        }
    }

    /// <summary>
    /// 昇順の文字位置を1始まりの論理行・列へ変換
    /// </summary>
    private static IReadOnlyList<SqlTextPosition> SqlTextPositions(
        string sql,
        IEnumerable<int> offsets)
    {
        var positions = new List<SqlTextPosition>();
        var currentOffset = 0;
        var line = 1;
        var column = 1;
        foreach (var targetOffset in offsets)
        {
            while (currentOffset < targetOffset)
            {
                if (sql[currentOffset] == '\r')
                {
                    if (currentOffset + 1 < sql.Length && sql[currentOffset + 1] == '\n')
                    {
                        currentOffset++;
                    }
                    line++;
                    column = 1;
                }
                else if (sql[currentOffset] == '\n')
                {
                    line++;
                    column = 1;
                }
                else
                {
                    column++;
                }
                currentOffset++;
            }
            positions.Add(new SqlTextPosition(line, column));
        }
        return positions;
    }

    /// <summary>
    /// SQLをScriptDomで解析
    /// </summary>
    private static TSqlScript? ParseScript(string sql, out IList<ParseError> errors)
    {
        var parser = new TSql160Parser(initialQuotedIdentifiers: false);
        using var reader = new StringReader(sql);
        return parser.Parse(reader, out errors) as TSqlScript;
    }

    /// <summary>
    /// SELECT式内の未修飾列を、変換定義から所属先が一意な場合だけSQL上のテーブル別名で修飾
    /// </summary>
    private static QualificationResult QualifyUnqualifiedSelectColumns(
        string sql,
        TSqlScript script,
        IReadOnlyList<MappingDefinition> mappings)
    {
        if (mappings.Count == 0)
        {
            return new QualificationResult(sql, []);
        }

        ColumnQualificationIndex? mappingIndex = null;
        var insertions = new Dictionary<int, QualificationInsertion>();
        foreach (var query in QuerySpecificationCollector.Collect(script))
        {
            var namedTables = query.FromClause?.TableReferences
                .SelectMany(EnumerateNamedTables)
                .ToArray() ?? [];
            if (namedTables.Length == 0)
            {
                continue;
            }

            foreach (var scalar in query.SelectElements.OfType<SelectScalarExpression>())
            {
                foreach (var column in ColumnReferenceCollector.Collect(scalar.Expression))
                {
                    var identifiers = column.MultiPartIdentifier?.Identifiers;
                    if (identifiers is null || identifiers.Count != 1)
                    {
                        continue;
                    }

                    mappingIndex ??= new ColumnQualificationIndex(mappings);
                    if (mappingIndex.IsEmpty)
                    {
                        return new QualificationResult(sql, []);
                    }

                    var qualifier = ResolveUniqueColumnQualifier(
                        sql,
                        identifiers[0].Value,
                        namedTables,
                        mappingIndex);
                    if (qualifier is not null)
                    {
                        var originalValue = FragmentText(sql, column);
                        insertions.TryAdd(
                            column.StartOffset,
                            new QualificationInsertion(
                                qualifier + ".",
                                new OutputReplacementQualification(
                                    column.StartLine,
                                    column.StartColumn,
                                    originalValue,
                                    qualifier + "." + originalValue)));
                    }
                }
            }
        }

        var orderedInsertions = insertions.OrderBy(item => item.Key).ToArray();
        return new QualificationResult(
            ApplyQualificationInsertions(sql, orderedInsertions),
            orderedInsertions
                .Select(item => item.Value.Qualification)
                .ToArray());
    }

    /// <summary>
    /// 元SQLのオフセットを保ったまま、全修飾子を1回の走査で挿入
    /// </summary>
    private static string ApplyQualificationInsertions(
        string sql,
        IReadOnlyList<KeyValuePair<int, QualificationInsertion>> insertions)
    {
        if (insertions.Count == 0)
        {
            return sql;
        }

        var addedLength = insertions.Sum(item => item.Value.Prefix.Length);
        var result = new StringBuilder(sql.Length + addedLength);
        var sourceOffset = 0;
        foreach (var insertion in insertions)
        {
            result.Append(sql, sourceOffset, insertion.Key - sourceOffset);
            result.Append(insertion.Value.Prefix);
            sourceOffset = insertion.Key;
        }
        result.Append(sql, sourceOffset, sql.Length - sourceOffset);
        return result.ToString();
    }

    /// <summary>
    /// 未修飾列に対応するFROMテーブルが変換定義上1つだけなら表示用修飾子を返す
    /// </summary>
    private static string? ResolveUniqueColumnQualifier(
        string sql,
        string fieldId,
        IReadOnlyList<NamedTableReference> namedTables,
        ColumnQualificationIndex mappingIndex)
    {
        var candidates = namedTables
            .Where(table => mappingIndex.AssociatesWithTable(fieldId, table))
            .DistinctBy(
                table => table.Alias?.Value ?? table.SchemaObject.BaseIdentifier.Value,
                StringComparer.OrdinalIgnoreCase)
            .Take(2)
            .ToArray();
        if (candidates.Length != 1)
        {
            return null;
        }

        var identifier = candidates[0].Alias ?? candidates[0].SchemaObject.BaseIdentifier;
        return FragmentText(sql, identifier);
    }

    private static OutputSheetPlan BuildStatement(
        string sql,
        TSqlStatement? statement,
        IReadOnlyList<MappingDefinition> mappings)
    {
        try
        {
            return statement switch
            {
                SelectStatement selectStatement => BuildSelectStatement(sql, selectStatement, mappings),
                InsertStatement insertStatement => BuildInsert(sql, insertStatement, mappings),
                UpdateStatement updateStatement => BuildUpdate(sql, updateStatement, mappings),
                DeleteStatement deleteStatement => BuildDelete(sql, deleteStatement, mappings),
                _ => CreateFallback(
                    sql,
                    "未対応のステートメント: " + StatementKind(statement),
                    statement)
            };
        }
        catch (UnsupportedOutputException ex)
        {
            return CreateFallback(sql, ex.Message, ex.Fragment);
        }
        catch (Exception ex)
        {
            return CreateFallback(sql, "解析結果の構成エラー: " + ex.Message);
        }
    }

    /// <summary>
    /// SELECTのサブクエリと全体クエリを出力順に構成
    /// </summary>
    private static OutputSheetPlan BuildSelectStatement(
        string sql,
        SelectStatement selectStatement,
        IReadOnlyList<MappingDefinition> mappings)
    {
        if (selectStatement.Into is not null &&
            UnwrapQueryExpression(selectStatement.QueryExpression) is QuerySpecification intoQuery)
        {
            return BuildSelectIntoStatement(sql, selectStatement, intoQuery, mappings);
        }

        var binaryAliases = CreateBinaryDisplayAliasPlan(selectStatement.QueryExpression);
        var (subqueries, plans) = BuildLeadingSubqueryPlans(
            sql,
            selectStatement,
            mappings,
            binaryAliases);

        var wholeChildren = DirectChildSubqueries(selectStatement.QueryExpression, subqueries);
        var wholePlan = BuildQueryExpression(
            sql,
            selectStatement.QueryExpression,
            mappings,
            "＜DB入出力項目定義＞",
            wholeChildren.Where(child => !child.IsNamed).Select(child => child.Name),
            binaryAliases);
        plans.Add(ReplaceSubqueries(wholePlan, sql, wholeChildren, binaryAliases));

        return CombinePlans(plans);
    }

    /// <summary>
    /// SELECT INTOを最上位SELECTとデータ移送表へ変換
    /// </summary>
    private static OutputSheetPlan BuildSelectIntoStatement(
        string sql,
        SelectStatement statement,
        QuerySpecification sourceQuery,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var (subqueries, plans) = BuildLeadingSubqueryPlans(sql, statement, mappings);

        var sourceChildren = DirectChildSubqueries(sourceQuery, subqueries);
        var sourcePlan = BuildSelect(
            sql,
            sourceQuery,
            mappings,
            "＜DB入出力項目定義＞",
            sourceChildren.Where(child => !child.IsNamed).Select(child => child.Name));
        plans.Add(ReplaceSubqueries(sourcePlan, sql, sourceChildren));

        var targetId = statement.Into.BaseIdentifier.Value;
        var targetColumns = BuildSelectIntoTargets(
            sql,
            sourceQuery,
            targetId,
            mappings);
        var targetName = ResolveTableName(targetId, mappings);
        var references = BuildTransferTableReferences(
            BuildTargetDisplay(targetName, targetId, targetId, includeIdentifier: false),
            BuildSelectTableDisplays(
                sourceQuery,
                mappings,
                sourceChildren.Where(child => !child.IsNamed).Select(child => child.Name)));
        var transfers = BuildSelectTransfers(sql, targetColumns, sourceQuery.SelectElements);
        var transferPlan = BuildDataTransferPlan(
            sql,
            references,
            transfers,
            null,
            null,
            mappings);
        plans.Add(ReplaceSubqueries(transferPlan, sql, sourceChildren));

        return CombinePlans(plans);
    }

    /// <summary>
    /// SQL断片内のサブクエリを内側から共通の描画計画へ変換
    /// </summary>
    private static (IReadOnlyList<SubqueryInfo> Subqueries, List<OutputSheetPlan> Plans)
        BuildLeadingSubqueryPlans(
            string sql,
            TSqlFragment fragment,
            IReadOnlyList<MappingDefinition> mappings,
            BinaryDisplayAliasPlan? outerBinaryAliases = null)
    {
        var subqueries = SubqueryCollector.Collect(fragment);
        var plans = new List<OutputSheetPlan>(subqueries.Count + 1);
        foreach (var subquery in subqueries)
        {
            var children = DirectChildSubqueries(subquery.QueryExpression, subqueries);
            var innerBinaryAliases = CreateBinaryDisplayAliasPlan(subquery.QueryExpression);
            OutputSheetPlan plan;
            using (PushDisplayAliases(
                outerBinaryAliases?.ContextContaining(subquery.QueryExpression)))
            {
                plan = BuildQueryExpression(
                    sql,
                    subquery.QueryExpression,
                    mappings,
                    $"サブクエリ[{subquery.Name}]",
                    children.Where(child => !child.IsNamed).Select(child => child.Name),
                    innerBinaryAliases);
                plan = ReplaceSubqueries(
                    plan,
                    sql,
                    children,
                    innerBinaryAliases);
            }
            plans.Add(plan);
        }

        return (subqueries, plans);
    }

    /// <summary>
    /// SELECT INTOの移送先項目名を変換定義から解決
    /// </summary>
    private static IReadOnlyList<string> BuildSelectIntoTargets(
        string sql,
        QuerySpecification sourceQuery,
        string targetId,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var columns = new List<string>(sourceQuery.SelectElements.Count);
        foreach (var element in sourceQuery.SelectElements)
        {
            if (element is SelectStarExpression)
            {
                columns.Add("全項目");
                continue;
            }

            if (element is not SelectScalarExpression scalar)
            {
                throw new UnsupportedOutputException(
                    "SELECT INTOの取得項目形式は未対応: " + element.GetType().Name,
                    element);
            }

            string fieldId;
            if (scalar.ColumnName is not null)
            {
                fieldId = FragmentText(sql, scalar.ColumnName);
            }
            else if (scalar.Expression is ColumnReferenceExpression column &&
                column.MultiPartIdentifier?.Identifiers.Count > 0)
            {
                fieldId = column.MultiPartIdentifier.Identifiers[^1].Value;
            }
            else
            {
                throw new UnsupportedOutputException(
                    "SELECT INTOの式には列エイリアスが必要",
                    scalar.Expression);
            }

            var fieldName = ResolveOutputFieldName(targetId, fieldId, mappings);
            columns.Add(fieldName);
        }

        return columns;
    }

    /// <summary>
    /// INSERT対象列を省略した場合にSELECT取得項目から仮の移送先項目名を構成
    /// </summary>
    private static IReadOnlyList<string> BuildInsertSelectTargets(
        string sql,
        InsertSpecification specification,
        QuerySpecification sourceQuery,
        IReadOnlyList<MappingDefinition> mappings)
    {
        if (specification.Columns.Count > 0)
        {
            return specification.Columns
                .Select(column => FragmentText(sql, column))
                .ToArray();
        }

        var targetId = specification.Target is NamedTableReference named
            ? named.SchemaObject.BaseIdentifier.Value
            : string.Empty;
        var targets = new List<string>(sourceQuery.SelectElements.Count);
        foreach (var element in sourceQuery.SelectElements)
        {
            if (element is SelectStarExpression star)
            {
                targets.Add(RenderSelectStar(DisplayText(sql, star)));
                continue;
            }

            if (element is not SelectScalarExpression scalar)
            {
                throw new UnsupportedOutputException(
                    "INSERT SELECTの取得項目形式は未対応: " + element.GetType().Name,
                    element);
            }

            string fieldId;
            if (scalar.ColumnName is not null)
            {
                fieldId = FragmentText(sql, scalar.ColumnName);
            }
            else if (scalar.Expression is ColumnReferenceExpression column &&
                column.MultiPartIdentifier?.Identifiers.Count > 0)
            {
                fieldId = column.MultiPartIdentifier.Identifiers[^1].Value;
            }
            else
            {
                fieldId = DisplayText(sql, scalar.Expression);
            }

            targets.Add(ResolveOutputFieldName(targetId, fieldId, mappings));
        }

        return targets;
    }

    /// <summary>
    /// SELECT INTOの出力別名を変換定義から和名へ解決
    /// </summary>
    private static string ResolveOutputFieldName(
        string targetId,
        string fieldId,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var physicalFieldId = mappings
            .FirstOrDefault(item =>
                item.ParserFieldId.Length > 0 &&
                string.Equals(item.ParserFieldId, fieldId, StringComparison.OrdinalIgnoreCase))
            ?.FieldId ?? fieldId;
        var mapping = mappings
            .Where(item => string.Equals(
                item.FieldId,
                physicalFieldId,
                StringComparison.OrdinalIgnoreCase))
            .Where(item => !string.IsNullOrWhiteSpace(item.FieldName))
            .OrderByDescending(item =>
                string.Equals(item.TableId, targetId, StringComparison.OrdinalIgnoreCase))
            .ThenByDescending(item => item.TableId == "-")
            .FirstOrDefault();
        return mapping?.FieldName ?? fieldId;
    }

    /// <summary>
    /// SELECT取得式を同じ位置の移送先項目へ対応付け
    /// </summary>
    private static IReadOnlyList<TransferItem> BuildSelectTransfers(
        string sql,
        IReadOnlyList<string> targets,
        IList<SelectElement> selectElements)
    {
        var transfers = new List<TransferItem>(targets.Count);
        for (var index = 0; index < targets.Count; index++)
        {
            if (selectElements[index] is SelectStarExpression star)
            {
                transfers.Add(new TransferItem(
                    targets[index],
                    RenderSelectStar(DisplayText(sql, star)),
                    string.Empty,
                    DisplayAliases: CurrentDisplayAliases.Value));
                continue;
            }

            if (selectElements[index] is not SelectScalarExpression scalar)
            {
                throw new UnsupportedOutputException(
                    "移送元として扱えない取得項目形式: " + selectElements[index].GetType().Name,
                    selectElements[index]);
            }

            transfers.Add(CreateCaseAwareTransferItem(
                sql,
                targets[index],
                DisplayText(sql, scalar.Expression),
                scalar.Expression));
        }

        return transfers;
    }

    /// <summary>
    /// INSERTの入力形式に応じた描画計画を作成
    /// </summary>
    private static OutputSheetPlan BuildInsert(
        string sql,
        InsertStatement statement,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var specification = statement.InsertSpecification;
        return specification.InsertSource switch
        {
            SelectInsertSource selectSource => BuildInsertSelect(
                sql,
                statement,
                specification,
                selectSource,
                mappings),
            ValuesInsertSource valuesSource => BuildInsertValues(
                sql,
                statement,
                specification,
                valuesSource,
                mappings),
            _ => CreateFallback(
                sql,
                "未対応のINSERT形式: " + InsertSourceKind(specification.InsertSource))
        };
    }

    /// <summary>
    /// INSERT SELECTを最上位SELECTとデータ移送表へ変換
    /// </summary>
    private static OutputSheetPlan BuildInsertSelect(
        string sql,
        InsertStatement statement,
        InsertSpecification specification,
        SelectInsertSource selectSource,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var sourceExpression = UnwrapQueryExpression(selectSource.Select);
        if (sourceExpression is BinaryQueryExpression binarySource)
        {
            return BuildInsertUnionSelect(
                sql,
                statement,
                specification,
                binarySource,
                mappings);
        }

        if (sourceExpression is not QuerySpecification sourceQuery)
        {
            return CreateFallback(sql, "未対応のINSERT形式: SELECTの集合演算", selectSource.Select);
        }

        if (specification.Columns.Count > 0 &&
            specification.Columns.Count != sourceQuery.SelectElements.Count)
        {
            return CreateFallback(
                sql,
                "INSERT SELECTの対象列数と取得項目数が一致しません",
                selectSource.Select);
        }

        var (subqueries, plans) = BuildLeadingSubqueryPlans(sql, statement, mappings);
        var sourceChildren = DirectChildSubqueries(sourceQuery, subqueries);
        var sourcePlan = BuildSelect(
            sql,
            sourceQuery,
            mappings,
            "＜DB入出力項目定義＞",
            sourceChildren.Where(child => !child.IsNamed).Select(child => child.Name));
        plans.Add(ReplaceSubqueries(sourcePlan, sql, sourceChildren));

        var targets = BuildInsertSelectTargets(sql, specification, sourceQuery, mappings);
        var transfers = BuildSelectTransfers(sql, targets, sourceQuery.SelectElements);
        var targetDisplay = BuildTargetTableDisplay(
            specification.Target,
            mappings,
            includeIdentifier: false);
        var references = BuildTransferTableReferences(
            targetDisplay,
            BuildSelectTableDisplays(
                sourceQuery,
                mappings,
                sourceChildren.Where(child => !child.IsNamed).Select(child => child.Name)));
        var transferPlan = BuildDataTransferPlan(
            sql,
            references,
            transfers,
            null,
            null,
            mappings);
        plans.Add(ReplaceSubqueries(transferPlan, sql, sourceChildren));

        return CombinePlans(plans);
    }

    /// <summary>
    /// INSERT SELECT内のUNIONを既存の複合SELECT表と分岐別のデータ移送表へ変換
    /// </summary>
    private static OutputSheetPlan BuildInsertUnionSelect(
        string sql,
        InsertStatement statement,
        InsertSpecification specification,
        BinaryQueryExpression sourceExpression,
        IReadOnlyList<MappingDefinition> mappings)
    {
        if (!ContainsOnlyUnionOperations(sourceExpression))
        {
            return CreateFallback(
                sql,
                "未対応のINSERT形式: UNION以外のSELECT集合演算",
                sourceExpression);
        }

        var branches = new List<QuerySpecification>();
        var separators = new List<string>();
        AddBinaryBranches(sourceExpression, branches, separators);
        if (branches.Count == 0 || separators.Count != branches.Count - 1)
        {
            return CreateFallback(
                sql,
                "INSERT SELECTのUNION分岐を取得できませんでした",
                sourceExpression);
        }

        var binaryAliases = BinaryDisplayAliasPlan.Create(branches);

        var targets = BuildInsertSelectTargets(sql, specification, branches[0], mappings);
        for (var index = 0; index < branches.Count; index++)
        {
            if (targets.Count != branches[index].SelectElements.Count)
            {
                return CreateFallback(
                    sql,
                    $"INSERT SELECTの対象列数と移送パターン{index + 1}の取得項目数が一致しません",
                    branches[index]);
            }
        }

        var (subqueries, plans) = BuildLeadingSubqueryPlans(
            sql,
            statement,
            mappings,
            binaryAliases);
        var sourceChildren = DirectChildSubqueries(sourceExpression, subqueries);
        var sourcePlan = BuildQueryExpression(
            sql,
            sourceExpression,
            mappings,
            "＜DB入出力項目定義＞",
            sourceChildren.Where(child => !child.IsNamed).Select(child => child.Name),
            binaryAliases);
        plans.Add(ReplaceSubqueries(
            sourcePlan,
            sql,
            sourceChildren,
            binaryAliases));

        var transferPatterns = new List<IReadOnlyList<TransferItem>>(branches.Count);
        foreach (var branch in branches)
        {
            using (PushDisplayAliases(binaryAliases.ContextFor(branch)))
            {
                transferPatterns.Add(BuildSelectTransfers(sql, targets, branch.SelectElements));
            }
        }
        var targetDisplay = BuildTargetTableDisplay(
            specification.Target,
            mappings,
            includeIdentifier: false);
        var references = BuildTransferTableReferences(
            targetDisplay,
            BuildBinaryTableDisplays(
                branches,
                mappings,
                sourceChildren.Where(child => !child.IsNamed).Select(child => child.Name),
                binaryAliases));
        var transferPlan = BuildLabeledDataTransferPlan(
            sql,
            references,
            transferPatterns,
            index => $"＜移送パターン{index + 1}＞");
        plans.Add(ReplaceSubqueries(transferPlan, sql, sourceChildren));

        return CombinePlans(plans);
    }

    /// <summary>
    /// 複合クエリがUNIONまたはUNION ALLだけで構成されているか判定
    /// </summary>
    private static bool ContainsOnlyUnionOperations(QueryExpression expression)
    {
        expression = UnwrapQueryExpression(expression);
        if (expression is QuerySpecification)
        {
            return true;
        }

        return expression is BinaryQueryExpression binary &&
            binary.BinaryQueryExpressionType == BinaryQueryExpressionType.Union &&
            ContainsOnlyUnionOperations(binary.FirstQueryExpression) &&
            ContainsOnlyUnionOperations(binary.SecondQueryExpression);
    }

    /// <summary>
    /// INSERT VALUESを行数に応じたデータ移送表へ変換
    /// </summary>
    private static OutputSheetPlan BuildInsertValues(
        string sql,
        InsertStatement statement,
        InsertSpecification specification,
        ValuesInsertSource valuesSource,
        IReadOnlyList<MappingDefinition> mappings)
    {
        if (valuesSource.IsDefaultValues)
        {
            return CreateFallback(sql, "未対応のINSERT形式: DEFAULT VALUES", valuesSource);
        }

        if (specification.Columns.Count == 0 || valuesSource.RowValues.Count == 0)
        {
            return CreateFallback(
                sql,
                "INSERT VALUESの対象列または値が指定されていません",
                valuesSource);
        }

        var transferRows = new List<IReadOnlyList<TransferItem>>(valuesSource.RowValues.Count);
        for (var rowIndex = 0; rowIndex < valuesSource.RowValues.Count; rowIndex++)
        {
            var values = valuesSource.RowValues[rowIndex].ColumnValues;
            if (specification.Columns.Count != values.Count)
            {
                var reason = valuesSource.RowValues.Count == 1
                    ? "INSERT VALUESの対象列数と値数が一致しません"
                    : $"INSERT VALUESの対象列数と{rowIndex + 1}行目の値数が一致しません";
                return CreateFallback(sql, reason, valuesSource.RowValues[rowIndex]);
            }

            var transfers = new List<TransferItem>(values.Count);
            for (var columnIndex = 0; columnIndex < values.Count; columnIndex++)
            {
                transfers.Add(CreateCaseAwareTransferItem(
                    sql,
                    FragmentText(sql, specification.Columns[columnIndex]),
                    FragmentText(sql, values[columnIndex]),
                    values[columnIndex]));
            }
            transferRows.Add(transfers);
        }

        var targetDisplay = BuildTargetTableDisplay(
            specification.Target,
            mappings,
            includeIdentifier: false);
        var (subqueries, plans) = BuildLeadingSubqueryPlans(sql, statement, mappings);
        var directChildren = DirectChildSubqueries(valuesSource, subqueries);
        var references = BuildTransferTableReferences(
            targetDisplay,
            directChildren.Select(child => child.Name));
        OutputSheetPlan transferPlan;
        if (transferRows.Count > 1)
        {
            transferPlan = BuildLabeledDataTransferPlan(
                sql,
                references,
                transferRows,
                index => $"＜VALUES {index + 1}行目＞");
        }
        else
        {
            transferPlan = BuildDataTransferPlan(
                sql,
                references,
                transferRows[0],
                null,
                null,
                mappings);
        }

        plans.Add(ReplaceSubqueries(transferPlan, sql, directChildren));
        return CombinePlans(plans);
    }

    /// <summary>
    /// DELETEの参照テーブルと条件をデータ移送表へ変換
    /// </summary>
    private static OutputSheetPlan BuildDelete(
        string sql,
        DeleteStatement statement,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var specification = statement.DeleteSpecification;
        var (subqueries, plans) = BuildLeadingSubqueryPlans(sql, statement, mappings);
        var directChildren = DirectChildSubqueries(specification, subqueries);
        var references = BuildModificationTableList(
            specification.Target,
            specification.FromClause,
            mappings,
            directChildren.Select(child => child.Name));

        var transferPlan = BuildDataTransferPlan(
            sql,
            references,
            [],
            specification.FromClause,
            specification.WhereClause,
            mappings);
        plans.Add(ReplaceSubqueries(transferPlan, sql, directChildren));
        return CombinePlans(plans);
    }

    /// <summary>
    /// UPDATEのSET、JOIN、WHEREをデータ移送表へ変換
    /// </summary>
    private static OutputSheetPlan BuildUpdate(
        string sql,
        UpdateStatement statement,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var specification = statement.UpdateSpecification;
        var (subqueries, plans) = BuildLeadingSubqueryPlans(sql, statement, mappings);
        var directChildren = DirectChildSubqueries(specification, subqueries);
        var references = BuildModificationTableList(
            specification.Target,
            specification.FromClause,
            mappings,
            directChildren.Select(child => child.Name));

        var transfers = specification.SetClauses
            .OfType<AssignmentSetClause>()
            .Where(clause => clause.Column is not null)
            .Select(clause => CreateUpdateTransferItem(
                sql,
                clause,
                directChildren,
                mappings))
            .ToArray();
        var transferPlan = BuildDataTransferPlan(
            sql,
            references,
            transfers,
            specification.FromClause,
            specification.WhereClause,
            mappings);
        plans.Add(ReplaceSubqueries(transferPlan, sql, directChildren));
        return CombinePlans(plans);
    }

    /// <summary>
    /// 更新系に共通するデータ移送表を構成
    /// </summary>
    private static OutputSheetPlan BuildDataTransferPlan(
        string sql,
        string references,
        IReadOnlyList<TransferItem> transfers,
        FromClause? fromClause,
        WhereClause? whereClause,
        IReadOnlyList<MappingDefinition> mappings,
        GroupByClause? groupByClause = null,
        HavingClause? havingClause = null,
        QuerySpecification? groupingQuery = null)
    {
        var cells = new List<OutputCell>
        {
            new(1, 1, "＜データ移送表＞"),
            new(2, 1, "参照テーブル: " + references)
        };
        var sections = new List<OutputSection>
        {
            new(OutputSectionKind.Reference, 2, 2)
        };
        var row = 3;

        if (transfers.Count > 0)
        {
            WriteTransferSection(cells, sections, sql, transfers, ref row);
        }

        WriteJoinSection(cells, sections, sql, fromClause, mappings, ref row);
        if (whereClause is not null)
        {
            WriteConditionSection(cells, sections, sql, "検索条件", whereClause.SearchCondition, ref row);
        }
        WriteGroupBySection(cells, sections, sql, groupByClause, groupingQuery, ref row);
        if (havingClause is not null)
        {
            WriteConditionSection(cells, sections, sql, "集計条件", havingClause.SearchCondition, ref row);
        }

        return new OutputSheetPlan(cells, sections, row - 1, false);
    }

    /// <summary>
    /// 複数の移送項目群をラベル付きの独立したデータ移送表として構成
    /// </summary>
    private static OutputSheetPlan BuildLabeledDataTransferPlan(
        string sql,
        string references,
        IReadOnlyList<IReadOnlyList<TransferItem>> transferPatterns,
        Func<int, string> buildLabel)
    {
        var cells = new List<OutputCell>
        {
            new(1, 1, "＜データ移送表＞"),
            new(2, 1, "参照テーブル: " + references)
        };
        var sections = new List<OutputSection>
        {
            new(OutputSectionKind.Reference, 2, 2)
        };
        var row = 3;

        for (var index = 0; index < transferPatterns.Count; index++)
        {
            cells.Add(new OutputCell(row, 1, buildLabel(index)));
            row++;
            WriteTransferSection(cells, sections, sql, transferPatterns[index], ref row);
        }

        return new OutputSheetPlan(cells, sections, row - 1, false);
    }

    /// <summary>
    /// データ移送表の見出し、項目、CASEグループを共通形式で追加
    /// </summary>
    private static void WriteTransferSection(
        ICollection<OutputCell> cells,
        ICollection<OutputSection> sections,
        string sql,
        IReadOnlyList<TransferItem> transfers,
        ref int row)
    {
        var startRow = row;
        var transferGroups = new List<OutputSection>();
        cells.Add(new OutputCell(row, 1, "項目"));
        cells.Add(new OutputCell(row, 19, "移送元"));
        cells.Add(new OutputCell(row, 37, "移送方法ほか"));
        row++;
        foreach (var transfer in transfers)
        {
            using var displayAliasScope = PushDisplayAliases(transfer.DisplayAliases);
            var itemStartRow = row;
            cells.Add(new OutputCell(row, 1, transfer.Target));
            if (transfer.Expression is not null &&
                transfer.DirectCases is { Count: > 0 })
            {
                if (transfer.RenderCaseInMethod && transfer.Source.Length > 0)
                {
                    cells.Add(new OutputCell(row, 19, transfer.Source));
                }
                var consumedRows = transfer.RenderCaseInMethod
                    ? WriteScalarExpression(
                        cells,
                        sql,
                        transfer.Expression,
                        row,
                        displayName: null,
                        valueColumn: 37,
                        markerColumn: 51,
                        detailColumn: 52,
                        directCases: transfer.DirectCases)
                    : WriteScalarExpression(
                        cells,
                        sql,
                        transfer.Expression,
                        row,
                        displayName: null,
                        valueColumn: 19,
                        markerColumn: 35,
                        detailColumn: 37,
                        directCases: transfer.DirectCases);
                row += consumedRows;
                if (transfer.RenderCaseInMethod && consumedRows > 1)
                {
                    transferGroups.Add(new OutputSection(
                        OutputSectionKind.TransferGroup,
                        itemStartRow,
                        row - 1));
                }
                continue;
            }

            if (transfer.Source.Length > 0)
            {
                cells.Add(new OutputCell(row, 19, transfer.Source));
            }
            if (transfer.Method.Length > 0)
            {
                cells.Add(new OutputCell(row, 37, transfer.Method));
            }
            row++;
        }
        sections.Add(new OutputSection(OutputSectionKind.Transfer, startRow, row - 1));
        foreach (var transferGroup in transferGroups)
        {
            sections.Add(transferGroup);
        }
    }

    /// <summary>
    /// 更新対象、FROM句、サブクエリを更新系の参照テーブル一覧へ統合
    /// </summary>
    private static string BuildModificationTableList(
        TableReference target,
        FromClause? fromClause,
        IReadOnlyList<MappingDefinition> mappings,
        IEnumerable<string> additionalTables)
    {
        var targetDisplay = BuildTargetTableDisplay(target, mappings, includeIdentifier: true);
        return BuildTransferTableReferences(
            targetDisplay,
            BuildTableDisplays(fromClause, mappings, additionalTables));
    }

    /// <summary>
    /// 更新対象テーブルの和名表示を作成
    /// </summary>
    private static string BuildTargetTableDisplay(
        TableReference target,
        IReadOnlyList<MappingDefinition> mappings,
        bool includeIdentifier)
    {
        if (target is not NamedTableReference named)
        {
            return MissingName;
        }

        var physicalTableId = named.SchemaObject.BaseIdentifier.Value;
        var tableId = named.Alias?.Value ?? physicalTableId;
        var tableName = ResolveTableName(named, mappings);
        return BuildTargetDisplay(
            tableName,
            tableId,
            physicalTableId,
            includeIdentifier);
    }

    /// <summary>
    /// 移送先の和名が未取得でも識別可能な物理テーブルIDを保持
    /// </summary>
    private static string BuildTargetDisplay(
        string tableName,
        string tableId,
        string physicalTableId,
        bool includeIdentifier)
    {
        if (tableName == MissingName)
        {
            return RegisterMissingTableDisplay(
                $"{MissingName}[{physicalTableId}]",
                physicalTableId,
                string.Empty);
        }

        return includeIdentifier ? $"{tableName}[{tableId}]" : tableName;
    }

    /// <summary>
    /// クエリ式の具象型に応じた描画計画を作成
    /// </summary>
    private static OutputSheetPlan BuildQueryExpression(
        string sql,
        QueryExpression expression,
        IReadOnlyList<MappingDefinition> mappings,
        string title,
        IEnumerable<string> additionalTables,
        BinaryDisplayAliasPlan? binaryAliases = null)
    {
        expression = UnwrapQueryExpression(expression);
        return expression switch
        {
            QuerySpecification query => BuildSelect(sql, query, mappings, title, additionalTables),
            BinaryQueryExpression binary => BuildBinaryQuery(
                sql,
                binary,
                mappings,
                title,
                additionalTables,
                binaryAliases),
            _ => CreateFallback(
                RawFragmentText(sql, expression),
                "未対応のクエリ式: " + expression.GetType().Name)
        };
    }

    /// <summary>
    /// UNIONなどの複合クエリを1フレームへ変換
    /// </summary>
    private static OutputSheetPlan BuildBinaryQuery(
        string sql,
        BinaryQueryExpression binary,
        IReadOnlyList<MappingDefinition> mappings,
        string title,
        IEnumerable<string> additionalTables,
        BinaryDisplayAliasPlan? binaryAliases = null)
    {
        var branches = new List<QuerySpecification>();
        var separators = new List<string>();
        AddBinaryBranches(binary, branches, separators);
        if (branches.Count == 0)
        {
            return CreateFallback(RawFragmentText(sql, binary), "複合クエリの分岐を取得できませんでした");
        }

        binaryAliases ??= BinaryDisplayAliasPlan.Create(branches);

        var tableDisplays = BuildBinaryTableDisplays(
            branches,
            mappings,
            additionalTables,
            binaryAliases);
        var cells = new List<OutputCell>
        {
            new(1, 1, title),
            new(2, 1, "参照テーブル: " + string.Join("、", tableDisplays))
        };
        var sections = new List<OutputSection>
        {
            new(OutputSectionKind.Reference, 2, 2)
        };
        var row = 3;

        for (var index = 0; index < branches.Count; index++)
        {
            OutputSheetPlan branchPlan;
            using (PushDisplayAliases(binaryAliases.ContextFor(branches[index])))
            {
                branchPlan = BuildSelect(sql, branches[index], mappings, title, []);
            }
            var bodyOffset = row - 3;
            cells.AddRange(branchPlan.Cells
                .Where(cell => cell.Row >= 3)
                .Select(cell => cell with { Row = cell.Row + bodyOffset }));
            sections.AddRange(branchPlan.Sections
                .Where(section => section.Kind != OutputSectionKind.Reference)
                .Select(section => section with
                {
                    StartRow = section.StartRow + bodyOffset,
                    EndRow = section.EndRow + bodyOffset
                }));
            row += Math.Max(0, branchPlan.RowCount - 2);

            if (index < separators.Count)
            {
                cells.Add(new OutputCell(row, 1, $"＜{separators[index]}＞"));
                sections.Add(new OutputSection(OutputSectionKind.Separator, row, row));
                row++;
            }
        }

        // UNION全体に付くORDER BY/OFFSETは、個別分岐ではなくフレーム末尾へ一度だけ出力する。
        using (PushDisplayAliases(binaryAliases.ContextFor(branches[0])))
        {
            WriteOffsetSection(cells, sections, sql, binary.OffsetClause, ref row);
            WriteOrderBySection(cells, sections, sql, binary.OrderByClause, ref row);
        }

        return new OutputSheetPlan(cells, sections, row - 1, false);
    }

    /// <summary>
    /// 複合クエリを左から分岐と演算子へ分解
    /// </summary>
    private static void AddBinaryBranches(
        QueryExpression expression,
        ICollection<QuerySpecification> branches,
        ICollection<string> separators)
    {
        expression = UnwrapQueryExpression(expression);
        if (expression is BinaryQueryExpression binary)
        {
            AddBinaryBranches(binary.FirstQueryExpression, branches, separators);
            separators.Add(BinaryOperatorText(binary));
            AddBinaryBranches(binary.SecondQueryExpression, branches, separators);
            return;
        }

        if (expression is QuerySpecification query)
        {
            branches.Add(query);
        }
    }

    /// <summary>
    /// 複合クエリなら全分岐で共有する表示用別名計画を作成
    /// </summary>
    private static BinaryDisplayAliasPlan? CreateBinaryDisplayAliasPlan(
        QueryExpression expression)
    {
        if (UnwrapQueryExpression(expression) is not BinaryQueryExpression binary)
        {
            return null;
        }

        var branches = new List<QuerySpecification>();
        var separators = new List<string>();
        AddBinaryBranches(binary, branches, separators);
        return branches.Count > 0
            ? BinaryDisplayAliasPlan.Create(branches)
            : null;
    }

    /// <summary>
    /// 複合クエリ演算子の表示文字列を取得
    /// </summary>
    private static string BinaryOperatorText(BinaryQueryExpression binary)
    {
        var operation = binary.BinaryQueryExpressionType switch
        {
            BinaryQueryExpressionType.Except => "EXCEPT",
            BinaryQueryExpressionType.Intersect => "INTERSECT",
            _ => "UNION"
        };
        return binary.All ? operation + " ALL" : operation;
    }

    /// <summary>
    /// 親クエリから直接参照されるサブクエリだけを取得
    /// </summary>
    private static IReadOnlyList<SubqueryInfo> DirectChildSubqueries(
        TSqlFragment parent,
        IReadOnlyList<SubqueryInfo> subqueries)
    {
        return subqueries
            .Where(candidate => ContainsFragment(parent, candidate.QueryExpression))
            .Where(candidate => !subqueries.Any(other =>
                !ReferenceEquals(candidate, other) &&
                ContainsFragment(parent, other.QueryExpression) &&
                ContainsFragment(other.QueryExpression, candidate.QueryExpression)))
            .ToArray();
    }

    /// <summary>
    /// AST断片が別の断片へ内包されるか判定
    /// </summary>
    private static bool ContainsFragment(TSqlFragment parent, TSqlFragment child)
    {
        if (parent.StartOffset == child.StartOffset && parent.FragmentLength == child.FragmentLength)
        {
            return false;
        }

        return child.StartOffset >= parent.StartOffset &&
            child.StartOffset + child.FragmentLength <= parent.StartOffset + parent.FragmentLength;
    }

    /// <summary>
    /// 条件式内のサブクエリ本文を出力名へ置換
    /// </summary>
    private static OutputSheetPlan ReplaceSubqueries(
        OutputSheetPlan plan,
        string sql,
        IReadOnlyList<SubqueryInfo> subqueries,
        BinaryDisplayAliasPlan? binaryAliases = null)
    {
        if (subqueries.Count == 0)
        {
            return plan;
        }

        var replacements = new List<SubqueryReplacement>(subqueries.Count);
        foreach (var subquery in subqueries)
        {
            using (PushDisplayAliases(
                binaryAliases?.ContextContaining(subquery.QueryExpression)))
            {
                replacements.Add(CreateSubqueryReplacement(sql, subquery));
            }
        }
        var cells = plan.Cells.Select(cell =>
        {
            var value = cell.Value;
            foreach (var replacement in replacements)
            {
                foreach (var sourceText in replacement.SourceTexts)
                {
                    value = value.Replace(sourceText, replacement.Name, StringComparison.Ordinal);
                }

                if (value.Contains(replacement.Name, StringComparison.Ordinal))
                {
                    value = replacement.ParenthesizedNamePattern.Replace(
                        value,
                        "(" + replacement.Name + ")");
                }
            }
            return value == cell.Value ? cell : cell with { Value = value };
        }).ToArray();
        return plan with { Cells = cells };
    }

    /// <summary>
    /// サブクエリ置換で全セルに共通するSQL表記と正規表現を事前計算
    /// </summary>
    private static SubqueryReplacement CreateSubqueryReplacement(
        string sql,
        SubqueryInfo subquery)
    {
        var sourceTexts = new[]
        {
            RawFragmentText(sql, subquery.QueryExpression),
            FragmentText(sql, subquery.QueryExpression),
            DisplayText(sql, subquery.QueryExpression)
        }
            .Where(text => text.Length > 0)
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        var parenthesizedNamePattern = new Regex(
            @"\(\s*" + Regex.Escape(subquery.Name) + @"\s*\)",
            RegexOptions.CultureInvariant);
        return new SubqueryReplacement(
            subquery.Name,
            sourceTexts,
            parenthesizedNamePattern);
    }

    /// <summary>
    /// フレーム間へ空行を置いて複数計画を連結
    /// </summary>
    private static OutputSheetPlan CombinePlans(IReadOnlyList<OutputSheetPlan> plans)
    {
        if (plans.Count == 1)
        {
            return plans[0];
        }

        var cells = new List<OutputCell>();
        var sections = new List<OutputSection>();
        var nextStartRow = 1;
        var lastRow = 0;
        int? fallbackQueryStartRow = null;
        int? fallbackQueryEndRow = null;
        for (var index = 0; index < plans.Count; index++)
        {
            var plan = plans[index];
            var offset = nextStartRow - 1;
            var shiftedFallbackStartRow = plan.FallbackQueryStartRow + offset;
            var shiftedFallbackEndRow = plan.FallbackQueryEndRow + offset;
            foreach (var cell in plan.Cells)
            {
                var shiftedCell = cell with { Row = cell.Row + offset };
                if (plan.IsFallback &&
                    cell.Row == plan.RowCount &&
                    cell.Column == 1 &&
                    plan.FallbackReason is not null &&
                    shiftedFallbackStartRow.HasValue &&
                    shiftedFallbackEndRow.HasValue)
                {
                    shiftedCell = shiftedCell with
                    {
                        Value = FormatFallbackMessage(
                            plan.FallbackReason,
                            shiftedFallbackStartRow.Value,
                            shiftedFallbackEndRow.Value)
                    };
                }

                cells.Add(shiftedCell);
            }
            sections.AddRange(plan.Sections.Select(section => section with
            {
                StartRow = section.StartRow + offset,
                EndRow = section.EndRow + offset
            }));
            if (!fallbackQueryStartRow.HasValue && shiftedFallbackStartRow.HasValue)
            {
                fallbackQueryStartRow = shiftedFallbackStartRow;
                fallbackQueryEndRow = shiftedFallbackEndRow;
            }
            lastRow = offset + plan.RowCount;
            nextStartRow = lastRow + 2;
        }

        return new OutputSheetPlan(
            cells,
            sections,
            lastRow,
            plans.Any(plan => plan.IsFallback),
            plans.FirstOrDefault(plan => plan.IsFallback)?.FallbackReason,
            fallbackQueryStartRow,
            fallbackQueryEndRow,
            FallbackSourceStartLine: plans
                .FirstOrDefault(plan => plan.IsFallback)?.FallbackSourceStartLine,
            FallbackSourceEndLine: plans
                .FirstOrDefault(plan => plan.IsFallback)?.FallbackSourceEndLine);
    }

    /// <summary>
    /// 単一SELECTの各句を仕様順に描画計画へ追加
    /// </summary>
    private static OutputSheetPlan BuildSelect(
        string sql,
        QuerySpecification query,
        IReadOnlyList<MappingDefinition> mappings,
        string title,
        IEnumerable<string> additionalTables)
    {
        var tableList = BuildTableList(query, mappings, additionalTables);
        var cells = new List<OutputCell>
        {
            new(1, 1, title),
            new(2, 1, "参照テーブル: " + tableList)
        };

        var sections = new List<OutputSection>
        {
            new(OutputSectionKind.Reference, 2, 2)
        };
        var row = 3;

        WriteOffsetSection(cells, sections, sql, query.OffsetClause, ref row);

        if (query.TopRowFilter is not null)
        {
            var startRow = row;
            cells.Add(new OutputCell(row, 1, "取得件数"));
            var topExpression = query.TopRowFilter.Expression is ParenthesisExpression parenthesized
                ? parenthesized.Expression
                : query.TopRowFilter.Expression;
            if (DirectCaseExpressions(topExpression).Count == 0)
            {
                cells.Add(new OutputCell(row, 7, RenderTopCount(sql, query.TopRowFilter.Expression)));
                row++;
            }
            else
            {
                row += WriteScalarExpression(
                    cells,
                    sql,
                    topExpression,
                    row,
                    displayName: null,
                    valueColumn: 7,
                    markerColumn: 15,
                    detailColumn: 17);
            }
            sections.Add(new OutputSection(OutputSectionKind.Standard, startRow, row - 1));
        }

        if (query.UniqueRowFilter == UniqueRowFilter.Distinct)
        {
            cells.Add(new OutputCell(row, 1, "重複除外"));
            cells.Add(new OutputCell(row, 7, "DISTINCT"));
            sections.Add(new OutputSection(OutputSectionKind.Standard, row, row));
            row++;
        }

        var itemStartRow = row;
        for (var index = 0; index < query.SelectElements.Count; index++)
        {
            row += WriteSelectElement(
                cells,
                sql,
                query,
                query.SelectElements[index],
                mappings,
                row,
                index + 1);
        }
        if (row > itemStartRow)
        {
            sections.Add(new OutputSection(OutputSectionKind.Standard, itemStartRow, row - 1));
        }

        WriteJoinSection(cells, sections, sql, query.FromClause, mappings, ref row);

        if (query.WhereClause is not null)
        {
            WriteConditionSection(cells, sections, sql, "検索条件", query.WhereClause.SearchCondition, ref row);
        }

        WriteGroupBySection(cells, sections, sql, query.GroupByClause, query, ref row);

        if (query.HavingClause is not null)
        {
            WriteConditionSection(cells, sections, sql, "集計条件", query.HavingClause.SearchCondition, ref row);
        }

        WriteOrderBySection(cells, sections, sql, query.OrderByClause, ref row);

        return new OutputSheetPlan(
            cells,
            sections,
            row - 1,
            false);
    }

    /// <summary>
    /// OFFSET/FETCHを帳票へ追加
    /// </summary>
    private static void WriteOffsetSection(
        ICollection<OutputCell> cells,
        ICollection<OutputSection> sections,
        string sql,
        OffsetClause? offsetClause,
        ref int row)
    {
        if (offsetClause is null)
        {
            return;
        }

        var startRow = row;
        cells.Add(new OutputCell(row, 1, "取得範囲"));
        var cases = DirectCaseExpressions(offsetClause);
        if (cases.Count == 0)
        {
            cells.Add(new OutputCell(
                row,
                7,
                DisplayText(sql, offsetClause, uppercaseOffsetKeywords: true)));
            row++;
        }
        else
        {
            cells.Add(new OutputCell(
                row,
                7,
                RenderExpressionWithCasePlaceholders(
                    sql,
                    offsetClause,
                    cases,
                    uppercaseOffsetKeywords: true)));
            cells.Add(new OutputCell(row, OffsetCaseMarkerColumn, "※"));
            row += WriteEmbeddedCaseBranches(
                cells,
                sql,
                cases,
                row,
                OffsetCaseDetailColumn);
        }

        sections.Add(new OutputSection(OutputSectionKind.Standard, startRow, row - 1));
    }

    /// <summary>
    /// ORDER BYを帳票へ追加
    /// </summary>
    private static void WriteOrderBySection(
        ICollection<OutputCell> cells,
        ICollection<OutputSection> sections,
        string sql,
        OrderByClause? orderByClause,
        ref int row)
    {
        if (orderByClause is null || orderByClause.OrderByElements.Count == 0)
        {
            return;
        }

        var startRow = row;
        for (var index = 0; index < orderByClause.OrderByElements.Count; index++)
        {
            var element = orderByClause.OrderByElements[index];
            if (index == 0)
            {
                cells.Add(new OutputCell(row, 1, "並び順"));
            }

            cells.Add(new OutputCell(row, 7, $"ソートキー{index + 1}"));
            cells.Add(new OutputCell(row, 15, ":"));
            var cases = DirectCaseExpressions(element.Expression);
            if (cases.Count > 0)
            {
                row += WriteScalarExpression(
                    cells,
                    sql,
                    element.Expression,
                    row,
                    displayName: null,
                    valueSuffix: element.SortOrder == SortOrder.Descending ? "(降順)" : string.Empty);
            }
            else
            {
                var value = DisplayText(sql, element.Expression);
                if (element.SortOrder == SortOrder.Descending)
                {
                    value += "(降順)";
                }

                cells.Add(new OutputCell(row, 17, value));
                row++;
            }
        }

        sections.Add(new OutputSection(OutputSectionKind.Standard, startRow, row - 1));
    }

    /// <summary>
    /// GROUP BY要素をSELECT系と更新系に共通のレイアウトで追加
    /// </summary>
    private static void WriteGroupBySection(
        ICollection<OutputCell> cells,
        ICollection<OutputSection> sections,
        string sql,
        GroupByClause? groupByClause,
        QuerySpecification? query,
        ref int row)
    {
        if (groupByClause is null)
        {
            return;
        }

        var startRow = row;
        for (var index = 0; index < groupByClause.GroupingSpecifications.Count; index++)
        {
            if (index == 0)
            {
                cells.Add(new OutputCell(row, 1, "グループ"));
            }

            cells.Add(new OutputCell(row, 7, $"グループキー{index + 1}"));
            cells.Add(new OutputCell(row, 15, ":"));
            var grouping = groupByClause.GroupingSpecifications[index];
            if (TryGetGroupingExpression(grouping, out var groupingExpression) &&
                DirectCaseExpressions(groupingExpression).Count > 0)
            {
                var alias = query is null
                    ? null
                    : FindSelectAlias(sql, query, groupingExpression);
                row += WriteScalarExpression(
                    cells,
                    sql,
                    groupingExpression,
                    row,
                    alias);
            }
            else
            {
                cells.Add(new OutputCell(row, 17, RenderGrouping(sql, grouping)));
                row++;
            }
        }
        sections.Add(new OutputSection(OutputSectionKind.Standard, startRow, row - 1));
    }

    /// <summary>
    /// 取得項目を出力し、消費した行数を返す
    /// </summary>
    private static int WriteSelectElement(
        ICollection<OutputCell> cells,
        string sql,
        QuerySpecification query,
        SelectElement element,
        IReadOnlyList<MappingDefinition> mappings,
        int row,
        int itemNumber)
    {
        cells.Add(new OutputCell(row, 7, $"取得項目{itemNumber}"));
        cells.Add(new OutputCell(row, 15, ":"));
        if (itemNumber == 1)
        {
            cells.Add(new OutputCell(row, 1, "取得項目"));
        }

        if (element is SelectScalarExpression scalar)
        {
            var alias = ResolveSelectElementDisplayName(sql, query, scalar, mappings);
            return WriteScalarExpression(cells, sql, scalar.Expression, row, alias);
        }

        cells.Add(new OutputCell(row, 17, RenderSelectElementForDisplay(sql, element)));
        return 1;
    }

    /// <summary>
    /// 式の別名が参照列の物理名と一致する場合は対応する列和名へ解決
    /// </summary>
    private static string? ResolveSelectElementDisplayName(
        string sql,
        QuerySpecification query,
        SelectScalarExpression scalar,
        IReadOnlyList<MappingDefinition> mappings)
    {
        if (scalar.ColumnName is null)
        {
            return null;
        }

        var alias = FragmentText(sql, scalar.ColumnName);
        var namedTables = query.FromClause?.TableReferences
            .SelectMany(EnumerateNamedTables)
            .ToArray() ?? [];
        if (DirectCaseExpressions(scalar.Expression).Count > 0)
        {
            return ResolveCaseAliasFieldName(
                scalar.Expression,
                alias,
                mappings,
                namedTables) ?? alias;
        }

        var columns = ColumnReferenceCollector.Collect(scalar.Expression);
        if (columns.Count == 0)
        {
            return alias;
        }

        var mapping = mappings
            .Where(item => string.Equals(
                item.FieldId,
                alias,
                StringComparison.OrdinalIgnoreCase))
            .Where(item =>
                !string.IsNullOrWhiteSpace(item.FieldName) &&
                item.FieldName != MissingName)
            .Where(item => columns.Any(column =>
                ColumnMatchesMapping(column, item, namedTables)))
            .OrderBy(item => item.TableId == "-")
            .FirstOrDefault();
        return mapping?.FieldName ?? alias;
    }

    /// <summary>
    /// CASEの全結果枝が同じ物理列名・和名へ解決できる場合だけ列和名を返す
    /// </summary>
    private static string? ResolveCaseAliasFieldName(
        ScalarExpression expression,
        string alias,
        IReadOnlyList<MappingDefinition> mappings,
        IReadOnlyList<NamedTableReference> namedTables)
    {
        var terminalColumns = new List<ColumnReferenceExpression>();
        if (!TryCollectCaseTerminalColumns(expression, terminalColumns))
        {
            return null;
        }

        string? commonFieldName = null;
        foreach (var column in terminalColumns)
        {
            if (!TryResolveCaseColumnFieldName(
                    column,
                    alias,
                    mappings,
                    namedTables,
                    out var fieldName))
            {
                return null;
            }

            if (commonFieldName is null)
            {
                commonFieldName = fieldName;
            }
            else if (!string.Equals(
                         commonFieldName,
                         fieldName,
                         StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }
        }

        return commonFieldName;
    }

    /// <summary>
    /// CASE条件を除外し、全THEN・ELSEの末端が単一列由来の場合だけ列を収集
    /// </summary>
    private static bool TryCollectCaseTerminalColumns(
        ScalarExpression expression,
        ICollection<ColumnReferenceExpression> terminalColumns)
    {
        var cases = DirectCaseExpressions(expression);
        if (cases.Count == 0)
        {
            var columns = ColumnReferenceCollector.Collect(expression);
            if (columns.Count != 1)
            {
                return false;
            }

            terminalColumns.Add(columns[0]);
            return true;
        }

        var expressionColumns = ColumnReferenceCollector.Collect(expression);
        if (expressionColumns.Any(column =>
                !cases.Any(caseExpression =>
                    ContainsFragment(caseExpression, column))))
        {
            return false;
        }

        foreach (var caseExpression in cases)
        {
            if (!TryCollectCaseBranchColumns(caseExpression, terminalColumns))
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// 検索CASE・単純CASEの全結果枝を再帰し、ELSE欠落を不成立とする
    /// </summary>
    private static bool TryCollectCaseBranchColumns(
        ScalarExpression expression,
        ICollection<ColumnReferenceExpression> terminalColumns)
    {
        switch (UnwrapScalarExpression(expression))
        {
            case SearchedCaseExpression searched:
                if (searched.ElseExpression is null)
                {
                    return false;
                }

                foreach (var clause in searched.WhenClauses)
                {
                    if (!TryCollectCaseTerminalColumns(
                            clause.ThenExpression,
                            terminalColumns))
                    {
                        return false;
                    }
                }

                return TryCollectCaseTerminalColumns(
                    searched.ElseExpression,
                    terminalColumns);

            case SimpleCaseExpression simple:
                if (simple.ElseExpression is null)
                {
                    return false;
                }

                foreach (var clause in simple.WhenClauses)
                {
                    if (!TryCollectCaseTerminalColumns(
                            clause.ThenExpression,
                            terminalColumns))
                    {
                        return false;
                    }
                }

                return TryCollectCaseTerminalColumns(
                    simple.ElseExpression,
                    terminalColumns);

            default:
                return false;
        }
    }

    /// <summary>
    /// CASE末端列を別名と同じ物理列の一意な和名へ解決
    /// </summary>
    private static bool TryResolveCaseColumnFieldName(
        ColumnReferenceExpression column,
        string alias,
        IReadOnlyList<MappingDefinition> mappings,
        IReadOnlyList<NamedTableReference> namedTables,
        out string fieldName)
    {
        var candidates = mappings
            .Where(item => string.Equals(
                item.FieldId,
                alias,
                StringComparison.OrdinalIgnoreCase))
            .Where(item =>
                !string.IsNullOrWhiteSpace(item.FieldName) &&
                item.FieldName != MissingName)
            .Where(item => ColumnMatchesMapping(column, item, namedTables))
            .ToArray();
        var tableSpecificCandidates = candidates
            .Where(item => item.TableId != "-")
            .ToArray();
        if (tableSpecificCandidates.Length > 0)
        {
            candidates = tableSpecificCandidates;
        }

        var fieldNames = candidates
            .Select(item => item.FieldName)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (fieldNames.Length != 1)
        {
            fieldName = string.Empty;
            return false;
        }

        fieldName = fieldNames[0];
        return true;
    }

    /// <summary>
    /// 列参照が和名定義の物理列と所属テーブルへ対応するか判定
    /// </summary>
    private static bool ColumnMatchesMapping(
        ColumnReferenceExpression column,
        MappingDefinition mapping,
        IReadOnlyList<NamedTableReference> namedTables)
    {
        var identifiers = column.MultiPartIdentifier?.Identifiers;
        if (identifiers is null || identifiers.Count == 0)
        {
            return false;
        }

        var fieldId = identifiers[^1].Value;
        if (!string.Equals(fieldId, mapping.FieldId, StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(fieldId, mapping.FieldName, StringComparison.OrdinalIgnoreCase) &&
            (mapping.ParserFieldId.Length == 0 ||
                !string.Equals(
                    fieldId,
                    mapping.ParserFieldId,
                    StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        if (mapping.TableId == "-")
        {
            return true;
        }

        if (identifiers.Count == 1)
        {
            return namedTables.Count == 1 &&
                MappingBelongsToTable(mapping, namedTables[0]);
        }

        var qualifier = identifiers[^2].Value;
        if (string.Equals(mapping.TableId, qualifier, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return namedTables.Any(table =>
            TableMatchesQualifier(table, qualifier) &&
            MappingBelongsToTable(mapping, table));
    }

    /// <summary>
    /// テーブル参照の物理名または別名が列修飾子と一致するか判定
    /// </summary>
    private static bool TableMatchesQualifier(
        NamedTableReference table,
        string qualifier)
    {
        return string.Equals(
                table.Alias?.Value,
                qualifier,
                StringComparison.OrdinalIgnoreCase) ||
            string.Equals(
                table.SchemaObject.BaseIdentifier.Value,
                qualifier,
                StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 和名定義がテーブル参照の物理名または別名へ所属するか判定
    /// </summary>
    private static bool MappingBelongsToTable(
        MappingDefinition mapping,
        NamedTableReference table)
    {
        return string.Equals(
                mapping.TableId,
                table.Alias?.Value,
                StringComparison.OrdinalIgnoreCase) ||
            string.Equals(
                mapping.TableId,
                table.SchemaObject.BaseIdentifier.Value,
                StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// スカラー式を表示名、CASE概要、分岐へ展開
    /// </summary>
    private static int WriteScalarExpression(
        ICollection<OutputCell> cells,
        string sql,
        ScalarExpression expression,
        int row,
        string? displayName,
        string valueSuffix = "",
        int valueColumn = 17,
        int markerColumn = 31,
        int detailColumn = 32,
        IReadOnlyList<ScalarExpression>? directCases = null)
    {
        var cases = directCases ?? DirectCaseExpressions(expression);
        if (cases.Count == 0)
        {
            if (displayName is null)
            {
                cells.Add(new OutputCell(row, valueColumn, DisplayText(sql, expression) + valueSuffix));
            }
            else
            {
                cells.Add(new OutputCell(row, valueColumn, displayName));
                cells.Add(new OutputCell(row, markerColumn, "※"));
                cells.Add(new OutputCell(row, detailColumn, DisplayText(sql, expression)));
            }

            return 1;
        }

        var isDirectCase = cases.Count == 1 && ReferenceEquals(expression, cases[0]);
        if (displayName is null)
        {
            var value = isDirectCase
                ? "CASE結果"
                : RenderExpressionWithCasePlaceholders(sql, expression, cases);
            cells.Add(new OutputCell(row, valueColumn, value + valueSuffix));
            cells.Add(new OutputCell(row, markerColumn, "※"));
            return isDirectCase
                ? WriteCaseBranches(cells, sql, cases[0], row, detailColumn)
                : WriteEmbeddedCaseBranches(cells, sql, cases, row, detailColumn);
        }

        cells.Add(new OutputCell(row, valueColumn, displayName));
        cells.Add(new OutputCell(row, markerColumn, "※"));
        if (isDirectCase)
        {
            return WriteCaseBranches(cells, sql, cases[0], row, detailColumn);
        }

        cells.Add(new OutputCell(
            row,
            detailColumn,
            RenderExpressionWithCasePlaceholders(sql, expression, cases)));
        return WriteEmbeddedCaseBranches(
            cells,
            sql,
            cases,
            row,
            detailColumn + (cases.Count == 1
                ? WrappedCaseBranchIndentColumns
                : WrappedMultipleCaseLabelIndentColumns));
    }

    /// <summary>
    /// 検索CASEのWHENとELSEを縦方向へ展開
    /// </summary>
    private static int WriteSearchedCase(
        ICollection<OutputCell> cells,
        string sql,
        SearchedCaseExpression expression,
        int startRow,
        int column = 32)
    {
        var row = startRow;
        foreach (var clause in expression.WhenClauses)
        {
            var conditions = BuildCaseConditionParts(clause.WhenExpression);
            for (var index = 0; index < conditions.Count - 1; index++)
            {
                var part = conditions[index];
                if (part.Connector.Length > 0)
                {
                    cells.Add(new OutputCell(
                        row,
                        column + part.ConnectorDepth * 2,
                        part.Connector));
                }
                cells.Add(new OutputCell(
                    row,
                    column + part.ExpressionDepth * 2,
                    CaseConditionDisplayText(sql, part)));
                row++;
            }

            var finalCondition = conditions[^1];
            if (finalCondition.Connector.Length > 0)
            {
                cells.Add(new OutputCell(
                    row,
                    column + finalCondition.ConnectorDepth * 2,
                    finalCondition.Connector));
            }
            row = WriteCaseResultLine(
                cells,
                sql,
                CaseConditionDisplayText(sql, finalCondition),
                clause.ThenExpression,
                row,
                column + finalCondition.ExpressionDepth * 2);
        }

        if (expression.ElseExpression is not null)
        {
            row = WriteCaseResultLine(
                cells,
                sql,
                "ELSE",
                expression.ElseExpression,
                row,
                column);
        }

        return Math.Max(1, row - startRow);
    }

    /// <summary>
    /// 単純CASEのWHENとELSEを縦方向へ展開
    /// </summary>
    private static int WriteSimpleCase(
        ICollection<OutputCell> cells,
        string sql,
        SimpleCaseExpression expression,
        int startRow,
        int column = 32)
    {
        var row = startRow;
        var input = DisplayText(sql, expression.InputExpression);
        foreach (var clause in expression.WhenClauses)
        {
            row = WriteCaseResultLine(
                cells,
                sql,
                $"{input} = {DisplayText(sql, clause.WhenExpression)}",
                clause.ThenExpression,
                row,
                column);
        }

        if (expression.ElseExpression is not null)
        {
            row = WriteCaseResultLine(
                cells,
                sql,
                "ELSE",
                expression.ElseExpression,
                row,
                column);
        }

        return Math.Max(1, row - startRow);
    }

    /// <summary>
    /// CASEのWHEN条件を括弧とAND/ORの論理構造に沿った表示部品へ変換
    /// </summary>
    private static IReadOnlyList<CaseConditionPart> BuildCaseConditionParts(
        BooleanExpression expression)
    {
        var parts = new List<CaseConditionPart>();
        AddCaseConditionParts(
            expression,
            groupDepth: 0,
            connector: string.Empty,
            connectorDepth: 0,
            parts);
        // CASE直下の先頭条件は、後続するAND/ORと同じ基準列へ戻す。
        if (parts.Count > 1 &&
            parts[0].Connector.Length == 0 &&
            parts[0].ExpressionDepth == 1)
        {
            parts[0] = parts[0] with { ExpressionDepth = 0 };
        }
        return parts;
    }

    /// <summary>
    /// ネストCASE用に、先頭の括弧グループと外側条件を同じ段落にし、
    /// 後続グループだけ1段ずつ深くした条件部品を作成
    /// </summary>
    private static IReadOnlyList<CaseConditionPart> BuildStructuredCaseConditionParts(
        BooleanExpression expression,
        bool isNestedCase)
    {
        var parts = new List<CaseConditionPart>();
        AddStructuredCaseConditionParts(
            expression,
            connector: string.Empty,
            connectorDepth: 0,
            operatorDepth: isNestedCase ? 1 : 0,
            isFirstPart: true,
            parts);
        return parts;
    }

    /// <summary>
    /// 論理二項式を左から展開し、演算子とオペランドの相対階層を記録
    /// </summary>
    private static void AddStructuredCaseConditionParts(
        BooleanExpression expression,
        string connector,
        int connectorDepth,
        int operatorDepth,
        bool isFirstPart,
        List<CaseConditionPart> parts)
    {
        var openingParentheses = 0;
        while (expression is BooleanParenthesisExpression parenthesized)
        {
            openingParentheses++;
            expression = parenthesized.Expression;
        }

        var firstPartIndex = parts.Count;
        if (expression is BooleanBinaryExpression binary)
        {
            var operands = CollectBooleanOperands(binary, binary.BinaryExpressionType);
            for (var index = 0; index < operands.Count; index++)
            {
                var operand = operands[index];
                var operandCore = UnwrapBooleanParentheses(operand);
                var operandIsGroup = operandCore is BooleanBinaryExpression;
                var childOperatorDepth = operatorDepth;
                if (operandIsGroup &&
                    (index > 0 || operand is not BooleanParenthesisExpression))
                {
                    childOperatorDepth++;
                }

                AddStructuredCaseConditionParts(
                    operand,
                    index == 0 ? connector : BooleanOperatorText(binary.BinaryExpressionType),
                    index == 0 ? connectorDepth : operatorDepth,
                    childOperatorDepth,
                    isFirstPart && index == 0,
                    parts);
            }
        }
        else
        {
            parts.Add(new CaseConditionPart(
                connector,
                connectorDepth,
                expression,
                isFirstPart ? 0 : connectorDepth + 1));
        }

        if (parts.Count <= firstPartIndex || openingParentheses == 0)
        {
            return;
        }

        parts[firstPartIndex] = parts[firstPartIndex] with
        {
            OpeningParentheses =
                parts[firstPartIndex].OpeningParentheses + openingParentheses
        };
        var lastPartIndex = parts.Count - 1;
        parts[lastPartIndex] = parts[lastPartIndex] with
        {
            ClosingParentheses =
                parts[lastPartIndex].ClosingParentheses + openingParentheses
        };
    }

    /// <summary>
    /// 条件式の外側括弧を除いた本体を取得
    /// </summary>
    private static BooleanExpression UnwrapBooleanParentheses(BooleanExpression expression)
    {
        while (expression is BooleanParenthesisExpression parenthesized)
        {
            expression = parenthesized.Expression;
        }

        return expression;
    }

    /// <summary>
    /// 条件本体へ元の論理グループを表す括弧を付けて表示
    /// </summary>
    private static string CaseConditionDisplayText(
        string sql,
        CaseConditionPart part)
    {
        return new string('(', part.OpeningParentheses) +
            ConditionDisplayText(sql, part.Expression) +
            new string(')', part.ClosingParentheses);
    }

    /// <summary>
    /// 同一階層の論理演算子を揃え、異なる論理グループを2列ずつ深くして追加
    /// </summary>
    private static void AddCaseConditionParts(
        BooleanExpression expression,
        int groupDepth,
        string connector,
        int connectorDepth,
        List<CaseConditionPart> parts)
    {
        if (expression is BooleanParenthesisExpression parenthesized)
        {
            var firstPartIndex = parts.Count;
            AddCaseConditionParts(
                parenthesized.Expression,
                groupDepth,
                connector,
                connectorDepth,
                parts);
            if (parts.Count > firstPartIndex)
            {
                parts[firstPartIndex] = parts[firstPartIndex] with
                {
                    OpeningParentheses = parts[firstPartIndex].OpeningParentheses + 1
                };
                var lastPartIndex = parts.Count - 1;
                parts[lastPartIndex] = parts[lastPartIndex] with
                {
                    ClosingParentheses = parts[lastPartIndex].ClosingParentheses + 1
                };
            }
            return;
        }

        if (expression is BooleanBinaryExpression binary)
        {
            var operands = CollectBooleanOperands(binary, binary.BinaryExpressionType);
            for (var index = 0; index < operands.Count; index++)
            {
                AddCaseConditionParts(
                    operands[index],
                    groupDepth + 1,
                    index == 0 ? connector : BooleanOperatorText(binary.BinaryExpressionType),
                    index == 0 ? connectorDepth : groupDepth,
                    parts);
            }
            return;
        }

        parts.Add(new CaseConditionPart(
            connector,
            connectorDepth,
            expression,
            groupDepth));
    }

    /// <summary>
    /// 括弧を越えず、連続する同種のANDまたはORを同一階層のオペランドへ展開
    /// </summary>
    private static IReadOnlyList<BooleanExpression> CollectBooleanOperands(
        BooleanExpression expression,
        BooleanBinaryExpressionType operatorType)
    {
        var operands = new List<BooleanExpression>();
        AddBooleanOperands(expression, operatorType, operands);
        return operands;
    }

    /// <summary>
    /// 同種の論理二項式を再帰的にオペランドへ追加
    /// </summary>
    private static void AddBooleanOperands(
        BooleanExpression expression,
        BooleanBinaryExpressionType operatorType,
        ICollection<BooleanExpression> operands)
    {
        if (expression is BooleanBinaryExpression binary &&
            binary.BinaryExpressionType == operatorType)
        {
            AddBooleanOperands(binary.FirstExpression, operatorType, operands);
            AddBooleanOperands(binary.SecondExpression, operatorType, operands);
            return;
        }

        operands.Add(expression);
    }

    /// <summary>
    /// 複合条件を持つ直接ネストCASEを、条件・結果・ELSEごとの行へ階層展開
    /// </summary>
    private static int WriteStructuredSearchedCase(
        ICollection<OutputCell> cells,
        string sql,
        SearchedCaseExpression expression,
        int startRow,
        int column,
        bool isNestedCase,
        string firstConditionPrefix = "")
    {
        var row = startRow;
        var prefixPending = firstConditionPrefix.Length > 0;
        foreach (var clause in expression.WhenClauses)
        {
            var conditions = BuildStructuredCaseConditionParts(
                clause.WhenExpression,
                isNestedCase);
            foreach (var part in conditions)
            {
                if (part.Connector.Length > 0)
                {
                    cells.Add(new OutputCell(
                        row,
                        column + part.ConnectorDepth * 2,
                        part.Connector));
                }

                var conditionText = CaseConditionDisplayText(sql, part);
                if (prefixPending)
                {
                    conditionText = firstConditionPrefix + conditionText;
                    prefixPending = false;
                }
                cells.Add(new OutputCell(
                    row,
                    column + part.ExpressionDepth * 2,
                    conditionText));
                row++;
            }

            row += WriteStructuredCaseResult(
                cells,
                sql,
                clause.ThenExpression,
                row,
                column,
                isNestedCase,
                hasElseLabel: false);
        }

        if (expression.ElseExpression is not null)
        {
            row += WriteStructuredCaseResult(
                cells,
                sql,
                expression.ElseExpression,
                row,
                column,
                isNestedCase,
                hasElseLabel: true);
        }

        return Math.Max(1, row - startRow);
    }

    /// <summary>
    /// 階層CASEのTHENまたはELSE結果を、子CASEとスカラ値で共通配置
    /// </summary>
    private static int WriteStructuredCaseResult(
        ICollection<OutputCell> cells,
        string sql,
        ScalarExpression result,
        int row,
        int column,
        bool isNestedCase,
        bool hasElseLabel)
    {
        var labelDepth = isNestedCase ? 1 : 0;
        if (hasElseLabel)
        {
            cells.Add(new OutputCell(row, column + labelDepth * 2, "ELSE"));
        }

        var cases = DirectCaseExpressions(result);
        var isDirectCase = cases.Count == 1 && ReferenceEquals(result, cases[0]);
        if (isDirectCase)
        {
            return WritePrefixedStructuredCase(
                cells,
                sql,
                cases[0],
                row,
                column + 2,
                "→ ");
        }

        var resultColumn = column + (labelDepth + 1) * 2;
        var resultText = cases.Count == 0
            ? DisplayText(sql, result)
            : RenderExpressionWithCasePlaceholders(sql, result, cases);
        cells.Add(new OutputCell(row, resultColumn, $"→ {resultText}"));
        if (cases.Count == 0)
        {
            return 1;
        }

        var consumedRows = WriteEmbeddedCaseBranches(
            cells,
            sql,
            cases,
            row + 1,
            resultColumn + 2);
        return consumedRows + 1;
    }

    /// <summary>
    /// 子CASEの先頭条件に親分岐からの矢印を付与して展開
    /// </summary>
    private static int WritePrefixedStructuredCase(
        ICollection<OutputCell> cells,
        string sql,
        ScalarExpression expression,
        int startRow,
        int column,
        string prefix)
    {
        if (expression is SearchedCaseExpression searchedCase)
        {
            return WriteStructuredSearchedCase(
                cells,
                sql,
                searchedCase,
                startRow,
                column,
                isNestedCase: true,
                firstConditionPrefix: prefix);
        }

        var nestedCells = new List<OutputCell>();
        var consumedRows = WriteCaseBranches(
            nestedCells,
            sql,
            expression,
            startRow,
            column);
        var first = nestedCells
            .OrderBy(cell => cell.Row)
            .ThenBy(cell => cell.Column)
            .FirstOrDefault();
        foreach (var cell in nestedCells)
        {
            cells.Add(ReferenceEquals(cell, first)
                ? cell with { Value = prefix + cell.Value }
                : cell);
        }
        return consumedRows;
    }

    /// <summary>
    /// 複合WHEN条件と直接ネストCASEの両方を持つか判定
    /// </summary>
    private static bool RequiresStructuredNestedCaseLayout(ScalarExpression expression)
    {
        return ContainsDirectNestedCaseBranch(expression) &&
            ContainsCompoundCaseCondition(expression);
    }

    /// <summary>
    /// CASEツリー内に直接CASEを返す分岐があるか判定
    /// </summary>
    private static bool ContainsDirectNestedCaseBranch(ScalarExpression expression)
    {
        IEnumerable<ScalarExpression> Results(ScalarExpression caseExpression)
        {
            return caseExpression switch
            {
                SearchedCaseExpression searched => searched.WhenClauses
                    .Select(clause => clause.ThenExpression)
                    .Concat(searched.ElseExpression is null
                        ? []
                        : [searched.ElseExpression]),
                SimpleCaseExpression simple => simple.WhenClauses
                    .Select(clause => clause.ThenExpression)
                    .Concat(simple.ElseExpression is null
                        ? []
                        : [simple.ElseExpression]),
                _ => []
            };
        }

        foreach (var result in Results(expression))
        {
            if (result is SearchedCaseExpression or SimpleCaseExpression)
            {
                return true;
            }

            foreach (var nestedCase in DirectCaseExpressions(result))
            {
                if (ContainsDirectNestedCaseBranch(nestedCase))
                {
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>
    /// CASEツリー内のWHENにANDまたはORの複合条件があるか判定
    /// </summary>
    private static bool ContainsCompoundCaseCondition(ScalarExpression expression)
    {
        switch (expression)
        {
            case SearchedCaseExpression searched:
                if (searched.WhenClauses.Any(clause =>
                    ContainsBooleanBinaryExpression(clause.WhenExpression)))
                {
                    return true;
                }
                return searched.WhenClauses
                        .SelectMany(clause => DirectCaseExpressions(clause.ThenExpression))
                        .Any(ContainsCompoundCaseCondition) ||
                    (searched.ElseExpression is not null &&
                        DirectCaseExpressions(searched.ElseExpression)
                            .Any(ContainsCompoundCaseCondition));
            case SimpleCaseExpression simple:
                return simple.WhenClauses
                        .SelectMany(clause => DirectCaseExpressions(clause.ThenExpression))
                        .Any(ContainsCompoundCaseCondition) ||
                    (simple.ElseExpression is not null &&
                        DirectCaseExpressions(simple.ElseExpression)
                            .Any(ContainsCompoundCaseCondition));
            default:
                return false;
        }
    }

    /// <summary>
    /// 括弧を辿った条件式内に論理二項式があるか判定
    /// </summary>
    private static bool ContainsBooleanBinaryExpression(BooleanExpression expression)
    {
        return expression switch
        {
            BooleanBinaryExpression => true,
            BooleanParenthesisExpression parenthesized =>
                ContainsBooleanBinaryExpression(parenthesized.Expression),
            BooleanNotExpression negated =>
                ContainsBooleanBinaryExpression(negated.Expression),
            _ => false
        };
    }

    /// <summary>
    /// CASE分岐の結果式と、その内側にあるCASEを階層表示
    /// </summary>
    private static int WriteCaseResultLine(
        ICollection<OutputCell> cells,
        string sql,
        string condition,
        ScalarExpression result,
        int row,
        int column)
    {
        var cases = DirectCaseExpressions(result);
        if (cases.Count == 0)
        {
            cells.Add(new OutputCell(row, column, $"{condition} → {DisplayText(sql, result)}"));
            return row + 1;
        }

        var isDirectCase = cases.Count == 1 && ReferenceEquals(result, cases[0]);
        if (isDirectCase)
        {
            return WriteDirectNestedCaseResult(
                cells,
                sql,
                condition,
                cases[0],
                row,
                column);
        }

        var resultText = RenderExpressionWithCasePlaceholders(sql, result, cases);
        cells.Add(new OutputCell(row, column, $"{condition} → {resultText}"));
        row++;
        var consumedRows = WriteEmbeddedCaseBranches(cells, sql, cases, row, column + 2);
        return row + consumedRows;
    }

    /// <summary>
    /// 直接ネストしたCASEは人工的なCASE行を作らず、親条件と最初の内側条件を同じ行へ連結
    /// </summary>
    private static int WriteDirectNestedCaseResult(
        ICollection<OutputCell> cells,
        string sql,
        string condition,
        ScalarExpression nestedCase,
        int row,
        int column)
    {
        var nestedCells = new List<OutputCell>();
        var consumedRows = WriteCaseBranches(
            nestedCells,
            sql,
            nestedCase,
            row,
            column + 2);
        if (nestedCells.Count == 0)
        {
            cells.Add(new OutputCell(row, column, $"{condition} → CASE"));
            return row + 1;
        }

        var first = nestedCells[0];
        cells.Add(new OutputCell(row, column, $"{condition} → {first.Value}"));
        foreach (var nestedCell in nestedCells.Skip(1))
        {
            cells.Add(nestedCell);
        }

        return row + consumedRows;
    }

    /// <summary>
    /// CASE種別に応じて分岐を縦方向へ展開
    /// </summary>
    private static int WriteCaseBranches(
        ICollection<OutputCell> cells,
        string sql,
        ScalarExpression expression,
        int startRow,
        int column = 32)
    {
        if (expression is SearchedCaseExpression structuredCase &&
            RequiresStructuredNestedCaseLayout(structuredCase))
        {
            return WriteStructuredSearchedCase(
                cells,
                sql,
                structuredCase,
                startRow,
                column,
                isNestedCase: false);
        }

        return expression switch
        {
            SearchedCaseExpression searchedCase =>
                WriteSearchedCase(cells, sql, searchedCase, startRow, column),
            SimpleCaseExpression simpleCase =>
                WriteSimpleCase(cells, sql, simpleCase, startRow, column),
            _ => 1
        };
    }

    /// <summary>
    /// 式に埋め込まれたCASEを単数または番号付きで展開
    /// </summary>
    private static int WriteEmbeddedCaseBranches(
        ICollection<OutputCell> cells,
        string sql,
        IReadOnlyList<ScalarExpression> cases,
        int startRow,
        int column)
    {
        if (cases.Count == 1)
        {
            return WriteCaseBranches(cells, sql, cases[0], startRow, column);
        }

        var row = startRow;
        for (var index = 0; index < cases.Count; index++)
        {
            cells.Add(new OutputCell(row, column, $"CASE結果{index + 1}"));
            row += WriteCaseBranches(
                cells,
                sql,
                cases[index],
                row,
                column + MultipleCaseBranchIndentColumns);
        }

        return Math.Max(1, row - startRow);
    }

    /// <summary>
    /// GROUP BY要素からスカラー式を取得
    /// </summary>
    private static bool TryGetGroupingExpression(
        GroupingSpecification grouping,
        out ScalarExpression expression)
    {
        if (grouping is ExpressionGroupingSpecification expressionGrouping)
        {
            expression = expressionGrouping.Expression;
            return true;
        }

        expression = null!;
        return false;
    }

    /// <summary>
    /// 式の直下にあるCASEを左から取得し、内側のCASEは親へ委ねる
    /// </summary>
    private static IReadOnlyList<ScalarExpression> DirectCaseExpressions(TSqlFragment expression)
    {
        var cases = CaseExpressionCollector.Collect(expression);
        return cases
            .Where(candidate => !cases.Any(other =>
                !ReferenceEquals(candidate, other) &&
                ContainsFragment(other, candidate)))
            .OrderBy(candidate => candidate.StartOffset)
            .ToArray();
    }

    /// <summary>
    /// 外側の式を保ったままCASE本文を結果名へ置換
    /// </summary>
    private static string RenderExpressionWithCasePlaceholders(
        string sql,
        TSqlFragment expression,
        IReadOnlyList<ScalarExpression> cases,
        bool uppercaseOffsetKeywords = false)
    {
        var value = DisplayText(
            sql,
            expression,
            uppercaseOffsetKeywords: uppercaseOffsetKeywords);
        for (var index = 0; index < cases.Count; index++)
        {
            var placeholder = cases.Count == 1
                ? "CASE結果"
                : $"CASE結果{index + 1}";
            var sourceTexts = new[]
            {
                DisplayText(sql, cases[index]),
                FragmentText(sql, cases[index]),
                RawFragmentText(sql, cases[index])
            };
            foreach (var sourceText in sourceTexts
                .Where(text => text.Length > 0)
                .Distinct(StringComparer.Ordinal)
                .OrderByDescending(text => text.Length))
            {
                value = value.Replace(sourceText, placeholder, StringComparison.Ordinal);
            }
        }

        return CompactSqlWhitespace(value);
    }

    /// <summary>
    /// リテラルと引用識別子の内部を維持し、不可視空白を半角スペース1個へ正規化
    /// </summary>
    private static string CompactSqlWhitespace(string value)
    {
        var result = new StringBuilder(value.Length);
        var pendingSpace = false;
        var inString = false;
        var inBracket = false;
        var inQuotedIdentifier = false;
        for (var index = 0; index < value.Length; index++)
        {
            var character = value[index];
            if (inString || inBracket || inQuotedIdentifier)
            {
                result.Append(character);
                if (inString && character == '\'' || inBracket && character == ']' ||
                    inQuotedIdentifier && character == '"')
                {
                    if (index + 1 < value.Length && value[index + 1] == character)
                    {
                        result.Append(value[++index]);
                    }
                    else
                    {
                        inString = false;
                        inBracket = false;
                        inQuotedIdentifier = false;
                    }
                }
                continue;
            }

            if (char.IsWhiteSpace(character))
            {
                pendingSpace = result.Length > 0;
                continue;
            }

            if (pendingSpace && result.Length > 0 &&
                result[^1] != '(' && character != ')' && character != ',' &&
                result[^1] != '.' && character != '.')
            {
                result.Append(' ');
            }
            pendingSpace = false;
            result.Append(character);
            inString = character == '\'';
            inBracket = character == '[';
            inQuotedIdentifier = character == '"';
        }

        return result.ToString().Trim();
    }

    /// <summary>
    /// SELECT項目と同じ式に付けられた列エイリアスを取得
    /// </summary>
    private static string? FindSelectAlias(
        string sql,
        QuerySpecification query,
        ScalarExpression expression)
    {
        var signature = ExpressionSignature(sql, expression);
        return query.SelectElements
            .OfType<SelectScalarExpression>()
            .Where(item => item.ColumnName is not null)
            .FirstOrDefault(item => ExpressionSignature(sql, item.Expression) == signature)
            ?.ColumnName is IdentifierOrValueExpression alias
                ? FragmentText(sql, alias)
                : null;
    }

    /// <summary>
    /// 空白と大文字小文字に依存しない式比較用文字列を作成
    /// </summary>
    private static string ExpressionSignature(string sql, TSqlFragment expression)
    {
        return Regex.Replace(
            RawFragmentText(sql, expression),
            @"\s+",
            string.Empty,
            RegexOptions.CultureInvariant).ToUpperInvariant();
    }

    /// <summary>
    /// WHEREやHAVINGを括弧構造に応じて行へ展開
    /// </summary>
    private static void WriteConditionSection(
        ICollection<OutputCell> cells,
        ICollection<OutputSection> sections,
        string sql,
        string label,
        BooleanExpression condition,
        ref int row)
    {
        var startRow = row;
        cells.Add(new OutputCell(row, 1, label));
        WriteConditionRows(cells, sql, condition, ref row);
        sections.Add(new OutputSection(OutputSectionKind.Standard, startRow, row - 1));
    }

    /// <summary>
    /// 論理条件を括弧構造に応じた共通配置で行へ展開
    /// </summary>
    private static void WriteConditionRows(
        ICollection<OutputCell> cells,
        string sql,
        BooleanExpression condition,
        ref int row)
    {
        var rootLayout = new ConditionLayout(7, 17, 17, 15);
        foreach (var part in FlattenBooleanExpression(condition))
        {
            row = WriteConditionPart(
                cells,
                sql,
                part.Expression,
                part.Connector,
                rootLayout,
                expandParentheses: true,
                row);
        }
    }

    /// <summary>
    /// 条件部品を括弧階層とCASE展開に応じたセルへ追加
    /// </summary>
    private static int WriteConditionPart(
        ICollection<OutputCell> cells,
        string sql,
        BooleanExpression expression,
        string connector,
        ConditionLayout layout,
        bool expandParentheses,
        int row)
    {
        if (TryGetParenthesizedCondition(expression, out var innerCondition, out var isNegated))
        {
            if (isNegated)
            {
                connector = connector.Length == 0 ? "NOT" : connector + " NOT";
            }

            var openColumn = connector.Length == 0
                ? layout.ConnectorColumn
                : layout.ConnectedGroupOpenColumn;
            if (expandParentheses && openColumn <= 17)
            {
                if (connector.Length > 0)
                {
                    cells.Add(new OutputCell(row, layout.ConnectorColumn, connector));
                }
                cells.Add(new OutputCell(row, openColumn, "("));

                var innerParts = FlattenBooleanExpression(innerCondition);
                var expandNested = innerParts.Any(part => part.Connector == "OR");
                var childLayout = CreateChildConditionLayout(
                    layout,
                    openColumn,
                    connector.Length > 0);
                foreach (var part in innerParts)
                {
                    row = WriteConditionPart(
                        cells,
                        sql,
                        part.Expression,
                        part.Connector,
                        childLayout,
                        expandNested,
                        row);
                }

                var closeColumn = connector.Length == 0
                    ? openColumn
                    : layout.ConnectorColumn;
                cells.Add(new OutputCell(row, closeColumn, ")"));
                return row + 1;
            }
        }

        if (connector.Length > 0)
        {
            cells.Add(new OutputCell(row, layout.ConnectorColumn, connector));
        }

        var valueColumn = connector.Length == 0
            ? layout.FirstValueColumn
            : layout.ConnectedValueColumn;
        if (TryWriteConditionCases(cells, sql, expression, valueColumn, row, out var consumedRows))
        {
            return row + consumedRows;
        }

        cells.Add(new OutputCell(row, valueColumn, ConditionDisplayText(sql, expression)));
        return row + 1;
    }

    /// <summary>
    /// 条件式の関数表記とA5M2由来の単項符号空白を正規化
    /// </summary>
    private static string ConditionDisplayText(string sql, TSqlFragment expression)
    {
        if (expression is ExistsPredicate existsPredicate)
        {
            return $"EXISTS ({DisplayText(sql, existsPredicate.Subquery.QueryExpression)})";
        }

        var rawText = RawFragmentText(sql, expression);
        var hasSpacedUnarySign = Regex.IsMatch(
            rawText,
            @"(?<![\w])([+-])\s+(?=\d)",
            RegexOptions.CultureInvariant);
        return DisplayText(
            sql,
            expression,
            uppercaseDateParts: hasSpacedUnarySign,
            compactUnarySigns: hasSpacedUnarySign);
    }

    /// <summary>
    /// 親の括弧位置から子条件の列配置を決定
    /// </summary>
    private static ConditionLayout CreateChildConditionLayout(
        ConditionLayout parent,
        int openColumn,
        bool hasConnector)
    {
        if (openColumn == 7)
        {
            return new ConditionLayout(15, 17, 17, 17);
        }

        if (hasConnector && openColumn - parent.ConnectorColumn >= 8)
        {
            return new ConditionLayout(openColumn, openColumn + 2, openColumn + 2, openColumn + 2);
        }

        return new ConditionLayout(openColumn + 2, openColumn + 2, openColumn + 4, openColumn + 4);
    }

    /// <summary>
    /// 条件式内のCASEを結果参照と分岐へ展開
    /// </summary>
    private static bool TryWriteConditionCases(
        ICollection<OutputCell> cells,
        string sql,
        BooleanExpression expression,
        int valueColumn,
        int row,
        out int consumedRows)
    {
        consumedRows = 0;
        var cases = DirectCaseExpressions(expression);
        if (cases.Count == 0)
        {
            return false;
        }

        cells.Add(new OutputCell(
            row,
            valueColumn,
            RenderExpressionWithCasePlaceholders(sql, expression, cases)));
        cells.Add(new OutputCell(row, 31, "※"));
        consumedRows = WriteEmbeddedCaseBranches(cells, sql, cases, row, 32);
        return true;
    }

    /// <summary>
    /// 条件から外側の括弧とNOTを取り出す
    /// </summary>
    private static bool TryGetParenthesizedCondition(
        BooleanExpression expression,
        out BooleanExpression innerCondition,
        out bool isNegated)
    {
        isNegated = false;
        if (expression is BooleanParenthesisExpression parenthesized)
        {
            innerCondition = parenthesized.Expression;
            return true;
        }

        if (expression is BooleanNotExpression negated &&
            negated.Expression is BooleanParenthesisExpression negatedParenthesis)
        {
            innerCondition = negatedParenthesis.Expression;
            isNegated = true;
            return true;
        }

        innerCondition = expression;
        return false;
    }

    /// <summary>
    /// JOINの組合せとON条件を出力
    /// </summary>
    private static void WriteJoinSection(
        ICollection<OutputCell> cells,
        ICollection<OutputSection> sections,
        string sql,
        FromClause? fromClause,
        IReadOnlyList<MappingDefinition> mappings,
        ref int row)
    {
        if (fromClause is null)
        {
            return;
        }

        var joins = fromClause.TableReferences.SelectMany(EnumerateJoins).ToArray();
        if (joins.Length == 0)
        {
            return;
        }

        var startRow = row;
        for (var joinIndex = 0; joinIndex < joins.Length; joinIndex++)
        {
            var join = joins[joinIndex];
            if (joinIndex == 0)
            {
                cells.Add(new OutputCell(row, 1, "結合条件"));
            }

            var leftTables = EnumerateJoinTables(join.FirstTableReference, mappings).ToArray();
            var rightTables = EnumerateJoinTables(join.SecondTableReference, mappings).ToArray();
            if (leftTables.Length == 0 || rightTables.Length == 0)
            {
                var unsupportedTable = leftTables.Length == 0
                    ? join.FirstTableReference
                    : join.SecondTableReference;
                throw new UnsupportedOutputException(
                    "JOIN対象のテーブル形式は未対応: " + unsupportedTable.GetType().Name,
                    unsupportedTable);
            }

            var referencedTableIds = DirectColumnQualifierCollector
                .Collect(join.SearchCondition)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var leftTable = BuildJoinSideDisplay(leftTables, referencedTableIds, useLastFallback: true);
            var rightTable = BuildJoinSideDisplay(rightTables, referencedTableIds, useLastFallback: false);
            var joinText = $"＜{leftTable} {JoinTypeText(join.QualifiedJoinType)} {rightTable}＞";
            cells.Add(new OutputCell(row, 17, joinText));
            row++;

            WriteConditionRows(cells, sql, join.SearchCondition, ref row);
        }

        sections.Add(new OutputSection(OutputSectionKind.Standard, startRow, row - 1));
    }

    /// <summary>
    /// 連鎖JOINを内側から列挙
    /// </summary>
    private static IEnumerable<QualifiedJoin> EnumerateJoins(TableReference table)
    {
        foreach (var child in EnumerateChildTableReferences(table))
        {
            foreach (var innerJoin in EnumerateJoins(child))
            {
                yield return innerJoin;
            }
        }

        if (table is QualifiedJoin join)
        {
            yield return join;
        }
    }

    /// <summary>
    /// JOIN片側の実テーブルと派生テーブル名を左から列挙
    /// </summary>
    private static IEnumerable<JoinTableDisplay> EnumerateJoinTables(
        TableReference table,
        IReadOnlyList<MappingDefinition> mappings)
    {
        switch (table)
        {
            case NamedTableReference named:
                yield return new JoinTableDisplay(
                    named.Alias?.Value ?? named.SchemaObject.BaseIdentifier.Value,
                    BuildTableDisplay(named, mappings));
                break;
            case QueryDerivedTable queryDerived:
                var queryDerivedId = queryDerived.Alias?.Value ?? MissingName;
                yield return new JoinTableDisplay(queryDerivedId, queryDerivedId);
                break;
            case InlineDerivedTable inlineDerived:
                var inlineDerivedId = inlineDerived.Alias?.Value ?? MissingName;
                yield return new JoinTableDisplay(
                    inlineDerivedId,
                    $"派生テーブル[{inlineDerivedId}]");
                break;
            case JoinTableReference:
                foreach (var child in EnumerateChildTableReferences(table))
                {
                    foreach (var display in EnumerateJoinTables(child, mappings))
                    {
                        yield return display;
                    }
                }
                break;
            case JoinParenthesisTableReference parenthesized:
                foreach (var display in EnumerateJoinTables(parenthesized.Join, mappings))
                {
                    yield return display;
                }
                break;
        }
    }

    /// <summary>
    /// テーブル参照ツリーの子要素を左から列挙
    /// </summary>
    private static IEnumerable<TableReference> EnumerateChildTableReferences(TableReference table)
    {
        switch (table)
        {
            case JoinTableReference join:
                yield return join.FirstTableReference;
                yield return join.SecondTableReference;
                break;
            case JoinParenthesisTableReference parenthesized:
                yield return parenthesized.Join;
                break;
        }
    }

    /// <summary>
    /// ON条件が参照するJOIN片側のテーブルを列挙し、解決できない場合は構造上の隣接テーブルへ戻す
    /// </summary>
    private static string BuildJoinSideDisplay(
        IReadOnlyList<JoinTableDisplay> tables,
        IReadOnlySet<string> referencedTableIds,
        bool useLastFallback)
    {
        var referencedDisplays = tables
            .Where(table => referencedTableIds.Contains(table.Identifier))
            .Select(table => table.Display)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (referencedDisplays.Length > 0)
        {
            return string.Join("、", referencedDisplays);
        }

        return useLastFallback ? tables[^1].Display : tables[0].Display;
    }

    /// <summary>
    /// JOIN種別の表示文字列を取得
    /// </summary>
    private static string JoinTypeText(QualifiedJoinType joinType)
    {
        return joinType switch
        {
            QualifiedJoinType.Inner => "INNER JOIN",
            QualifiedJoinType.LeftOuter => "LEFT JOIN",
            QualifiedJoinType.RightOuter => "RIGHT JOIN",
            QualifiedJoinType.FullOuter => "FULL JOIN",
            _ => "JOIN"
        };
    }

    /// <summary>
    /// 直下のANDとORを条件部品へ分解
    /// </summary>
    private static IReadOnlyList<ConditionPart> FlattenBooleanExpression(BooleanExpression expression)
    {
        var parts = new List<ConditionPart>();
        AddBooleanParts(expression, string.Empty, parts);
        return parts;
    }

    /// <summary>
    /// 論理二項式を再帰的に条件部品へ追加
    /// </summary>
    private static void AddBooleanParts(
        BooleanExpression expression,
        string connector,
        ICollection<ConditionPart> parts)
    {
        if (expression is BooleanBinaryExpression binary)
        {
            AddBooleanParts(binary.FirstExpression, connector, parts);
            AddBooleanParts(binary.SecondExpression, BooleanOperatorText(binary.BinaryExpressionType), parts);
            return;
        }

        parts.Add(new ConditionPart(connector, expression));
    }

    /// <summary>
    /// 論理演算子の表示文字列を取得
    /// </summary>
    private static string BooleanOperatorText(BooleanBinaryExpressionType operatorType)
    {
        return operatorType == BooleanBinaryExpressionType.Or ? "OR" : "AND";
    }

    /// <summary>
    /// GROUP BY要素を表示文字列へ変換
    /// </summary>
    private static string RenderGrouping(string sql, GroupingSpecification grouping)
    {
        return grouping is ExpressionGroupingSpecification expressionGrouping
            ? DisplayText(sql, expressionGrouping.Expression)
            : DisplayText(sql, grouping);
    }

    /// <summary>
    /// TOP件数から外側の括弧を除いて表示
    /// </summary>
    private static string RenderTopCount(string sql, ScalarExpression expression)
    {
        return expression is ParenthesisExpression parenthesized
            ? FragmentText(sql, parenthesized.Expression)
            : FragmentText(sql, expression);
    }

    /// <summary>
    /// SELECTの参照テーブル一覧を作成
    /// </summary>
    private static string BuildTableList(
        QuerySpecification query,
        IReadOnlyList<MappingDefinition> mappings,
        IEnumerable<string> additionalTables)
    {
        var displays = BuildSelectTableDisplays(query, mappings, additionalTables);
        return displays.Count == 0 ? "なし" : string.Join("、", displays);
    }

    /// <summary>
    /// SELECTの参照テーブルを表示順・重複なしで列挙
    /// </summary>
    private static IReadOnlyList<string> BuildSelectTableDisplays(
        QuerySpecification query,
        IReadOnlyList<MappingDefinition> mappings,
        IEnumerable<string> additionalTables)
    {
        var localIdentifiers = (query.FromClause?.TableReferences
            .SelectMany(EnumerateTableIdentifiers)
            ?? [])
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var correlatedDisplays = DirectColumnQualifierCollector.Collect(query)
            .Where(identifier => !localIdentifiers.Contains(identifier))
            .Select(identifier => BuildTableDisplay(identifier, mappings));
        var displays = correlatedDisplays
            .Concat(BuildTableDisplays(query.FromClause, mappings, additionalTables))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        return displays;
    }

    /// <summary>
    /// UNION各分岐の参照テーブルを表示順・重複なしで列挙
    /// </summary>
    private static IReadOnlyList<string> BuildBinaryTableDisplays(
        IEnumerable<QuerySpecification> branches,
        IReadOnlyList<MappingDefinition> mappings,
        IEnumerable<string> additionalTables,
        BinaryDisplayAliasPlan? binaryAliases = null)
    {
        var branchArray = branches.ToArray();
        binaryAliases ??= BinaryDisplayAliasPlan.Create(branchArray);
        var displays = new List<string>();
        foreach (var branch in branchArray)
        {
            using (PushDisplayAliases(binaryAliases.ContextFor(branch)))
            {
                displays.AddRange(BuildTableDisplays(branch.FromClause, mappings, []));
            }
        }

        return displays
            .Concat(additionalTables)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    /// <summary>
    /// データ移送表の移送先と移送元テーブルを表示順・重複なしで統合
    /// </summary>
    private static string BuildTransferTableReferences(
        string targetDisplay,
        IEnumerable<string> sourceDisplays)
    {
        return string.Join(
            "、",
            new[] { targetDisplay }
                .Concat(sourceDisplays)
                .Where(display => display.Length > 0 && display != "なし")
                .Distinct(StringComparer.OrdinalIgnoreCase));
    }

    /// <summary>
    /// FROM句を重複のないテーブル表示へ変換
    /// </summary>
    private static IReadOnlyList<string> BuildTableDisplays(
        FromClause? fromClause,
        IReadOnlyList<MappingDefinition> mappings,
        IEnumerable<string> additionalTables)
    {
        var allowStandaloneTableName = fromClause?.TableReferences
            .SelectMany(EnumerateNamedTables)
            .Take(2)
            .Count() == 1;
        return (fromClause?.TableReferences
            .SelectMany(table => EnumerateTableDisplays(
                table,
                mappings,
                allowStandaloneTableName))
            ?? [])
            .Concat(additionalTables)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    /// <summary>
    /// テーブル参照ツリーを帳票上の参照名へ変換
    /// </summary>
    private static IEnumerable<string> EnumerateTableDisplays(
        TableReference table,
        IReadOnlyList<MappingDefinition> mappings,
        bool allowStandaloneTableName)
    {
        switch (table)
        {
            case NamedTableReference named:
                yield return BuildTableDisplay(named, mappings, allowStandaloneTableName);
                break;
            case JoinTableReference:
                foreach (var child in EnumerateChildTableReferences(table))
                {
                    foreach (var display in EnumerateTableDisplays(
                        child,
                        mappings,
                        allowStandaloneTableName))
                    {
                        yield return display;
                    }
                }
                break;
            case JoinParenthesisTableReference parenthesized:
                foreach (var display in EnumerateTableDisplays(
                    parenthesized.Join,
                    mappings,
                    allowStandaloneTableName))
                {
                    yield return display;
                }
                break;
            case InlineDerivedTable inlineDerived:
                yield return $"派生テーブル[{inlineDerived.Alias?.Value ?? MissingName}]";
                break;
            case QueryDerivedTable queryDerived:
                yield return queryDerived.Alias?.Value ?? MissingName;
                break;
        }
    }

    /// <summary>
    /// テーブル参照を和名と識別子の表示へ変換
    /// </summary>
    private static string BuildTableDisplay(
        NamedTableReference table,
        IReadOnlyList<MappingDefinition> mappings,
        bool allowStandaloneTableName = false)
    {
        var physicalTableId = table.SchemaObject.BaseIdentifier.Value;
        var sourceTableId = table.Alias?.Value ?? physicalTableId;
        var tableId = CurrentDisplayAliases.Value?.DisplayAliasFor(table) ?? sourceTableId;
        var tableName = ResolveDisplayTableName(table, sourceTableId, tableId, mappings);
        if (tableName == MissingName && allowStandaloneTableName)
        {
            var standaloneTableName = ResolveStandaloneTableName(mappings);
            if (standaloneTableName is not null)
            {
                return table.Alias is null
                    ? standaloneTableName
                    : $"{standaloneTableName}[{tableId}]";
            }
        }

        if (tableName == MissingName &&
            table.Alias is not null &&
            CurrentDisplayAliases.Value?.ShouldPreservePhysicalTableId(table) == true)
        {
            return RegisterMissingTableDisplay(
                $"{MissingName}[{physicalTableId}][{tableId}]",
                physicalTableId,
                $"[{tableId}]");
        }

        var display = $"{tableName}[{tableId}]";
        if (tableName != MissingName)
        {
            return display;
        }

        return RegisterMissingTableDisplay(
            display,
            physicalTableId,
            table.Alias is null ? string.Empty : $"[{tableId}]");
    }

    /// <summary>
    /// 元のSQL別名または物理名からテーブル和名を解決
    /// </summary>
    private static string ResolveDisplayTableName(
        NamedTableReference table,
        string sourceTableId,
        string displayTableId,
        IReadOnlyList<MappingDefinition> mappings)
    {
        if (!string.Equals(sourceTableId, displayTableId, StringComparison.OrdinalIgnoreCase))
        {
            return ResolveTableName(
                table.SchemaObject.BaseIdentifier.Value,
                mappings);
        }

        if (CurrentDisplayAliases.Value?.ShouldPreservePhysicalTableId(table) == true)
        {
            var physicalName = ResolveTableName(
                table.SchemaObject.BaseIdentifier.Value,
                mappings);
            if (physicalName != MissingName)
            {
                return physicalName;
            }
        }

        var sourceName = ResolveTableName(table, mappings);
        if (sourceName != MissingName ||
            CurrentDisplayAliases.Value?.ShouldPreservePhysicalTableId(table) != true)
        {
            return sourceName;
        }

        return ResolveTableName(table.SchemaObject.BaseIdentifier.Value, mappings);
    }

    /// <summary>
    /// テーブル識別子から和名付き表示を作成
    /// </summary>
    private static string BuildTableDisplay(
        string tableId,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var aliases = CurrentDisplayAliases.Value;
        var displayTableId = aliases?.DisplayAliasFor(tableId) ?? tableId;
        var physicalTableId = aliases?.PhysicalTableIdFor(tableId);
        var usesSyntheticAlias = !string.Equals(
            tableId,
            displayTableId,
            StringComparison.OrdinalIgnoreCase);
        var tableName = usesSyntheticAlias && physicalTableId is not null
            ? ResolveTableName(physicalTableId, mappings)
            : ResolveTableName(tableId, mappings);
        if (tableName == MissingName &&
            physicalTableId is not null &&
            aliases?.ShouldPreservePhysicalTableId(tableId) == true)
        {
            tableName = ResolveTableName(physicalTableId, mappings);
        }
        if (tableName == MissingName &&
            physicalTableId is not null &&
            aliases?.ShouldPreservePhysicalTableId(tableId) == true)
        {
            return RegisterMissingTableDisplay(
                $"{MissingName}[{physicalTableId}][{displayTableId}]",
                physicalTableId,
                $"[{displayTableId}]");
        }
        var display = $"{tableName}[{displayTableId}]";
        return tableName == MissingName && physicalTableId is not null
            ? RegisterMissingTableDisplay(
                display,
                physicalTableId,
                $"[{displayTableId}]")
            : display;
    }

    /// <summary>
    /// 物理テーブル名を優先し、未解決の場合だけ元の別名で和名を検索
    /// </summary>
    private static string ResolveTableName(
        NamedTableReference table,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var baseTableId = table.SchemaObject.BaseIdentifier.Value;
        var displayTableId = table.Alias?.Value ?? baseTableId;
        var physicalNames = ResolveTableNames(baseTableId, mappings);
        if (physicalNames.Count == 1)
        {
            return physicalNames[0];
        }
        if (physicalNames.Count > 1)
        {
            return MissingName;
        }

        var displayNames = ResolveTableNames(displayTableId, mappings);
        return displayNames.Count == 1 ? displayNames[0] : MissingName;
    }

    /// <summary>
    /// テーブルIDに紐づく一意な和名候補を列挙
    /// </summary>
    private static IReadOnlyList<string> ResolveTableNames(
        string tableId,
        IReadOnlyList<MappingDefinition> mappings)
    {
        return mappings
            .Where(mapping => string.Equals(mapping.TableId, tableId, StringComparison.OrdinalIgnoreCase))
            .Select(mapping => mapping.TableName.Trim())
            .Where(tableName => !string.IsNullOrWhiteSpace(tableName))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(2)
            .ToArray();
    }

    /// <summary>
    /// テーブルIDから和名を解決
    /// </summary>
    private static string ResolveTableName(
        string tableId,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var tableNames = ResolveTableNames(tableId, mappings);
        return tableNames.Count == 1 ? tableNames[0] : MissingName;
    }

    /// <summary>
    /// 単独フィールド定義から一意なテーブル和名を解決
    /// </summary>
    private static string? ResolveStandaloneTableName(IReadOnlyList<MappingDefinition> mappings)
    {
        var tableNames = mappings
            .Where(mapping => string.Equals(
                mapping.TableId.Trim(),
                "-",
                StringComparison.Ordinal))
            .Select(mapping => mapping.TableName.Trim())
            .Where(tableName =>
                tableName.Length > 0 &&
                tableName != "-" &&
                !tableName.Contains(MissingName, StringComparison.Ordinal))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(2)
            .ToArray();
        return tableNames.Length == 1 ? tableNames[0] : null;
    }

    /// <summary>
    /// テーブル参照ツリーから実テーブルを左から列挙
    /// </summary>
    private static IEnumerable<NamedTableReference> EnumerateNamedTables(TableReference table)
    {
        switch (table)
        {
            case NamedTableReference named:
                yield return named;
                break;
            case JoinTableReference:
                foreach (var child in EnumerateChildTableReferences(table))
                {
                    foreach (var item in EnumerateNamedTables(child))
                    {
                        yield return item;
                    }
                }
                break;
            case JoinParenthesisTableReference parenthesized:
                foreach (var item in EnumerateNamedTables(parenthesized.Join))
                {
                    yield return item;
                }
                break;
        }
    }

    /// <summary>
    /// FROM句のローカルテーブル識別子を列挙
    /// </summary>
    private static IEnumerable<string> EnumerateTableIdentifiers(TableReference table)
    {
        switch (table)
        {
            case NamedTableReference named:
                yield return named.Alias?.Value ?? named.SchemaObject.BaseIdentifier.Value;
                break;
            case QueryDerivedTable queryDerived when queryDerived.Alias is not null:
                yield return queryDerived.Alias.Value;
                break;
            case InlineDerivedTable inlineDerived when inlineDerived.Alias is not null:
                yield return inlineDerived.Alias.Value;
                break;
            case JoinTableReference:
                foreach (var child in EnumerateChildTableReferences(table))
                {
                    foreach (var identifier in EnumerateTableIdentifiers(child))
                    {
                        yield return identifier;
                    }
                }
                break;
            case JoinParenthesisTableReference parenthesized:
                foreach (var identifier in EnumerateTableIdentifiers(parenthesized.Join))
                {
                    yield return identifier;
                }
                break;
        }
    }

    /// <summary>
    /// SELECT取得項目を帳票向け表記で表示
    /// </summary>
    private static string RenderSelectElementForDisplay(string sql, SelectElement element)
    {
        return element switch
        {
            SelectScalarExpression scalar => DisplayText(sql, scalar.Expression),
            SelectStarExpression star => RenderSelectStar(DisplayText(sql, star)),
            _ => DisplayText(sql, element)
        };
    }

    /// <summary>
    /// SELECTのアスタリスクを帳票用の全項目表記へ変換
    /// </summary>
    private static string RenderSelectStar(string starText)
    {
        if (starText == "*")
        {
            return "全項目";
        }

        return starText.EndsWith(".*", StringComparison.Ordinal)
            ? starText[..^1] + "全項目"
            : starText;
    }

    /// <summary>
    /// 現在の描画スコープへ表示用別名を追加し、破棄時に元へ戻す
    /// </summary>
    private static IDisposable PushDisplayAliases(DisplayAliasContext? aliases)
    {
        var previous = CurrentDisplayAliases.Value;
        CurrentDisplayAliases.Value = DisplayAliasContext.Combine(previous, aliases);
        return new DisplayAliasScope(previous);
    }

    /// <summary>
    /// クエリ式を囲む括弧ノードを除去
    /// </summary>
    private static QueryExpression UnwrapQueryExpression(QueryExpression expression)
    {
        while (expression is QueryParenthesisExpression parenthesized)
        {
            expression = parenthesized.QueryExpression;
        }

        return expression;
    }

    /// <summary>
    /// AST位置から元SQLの文字列を取得
    /// </summary>
    private static string FragmentText(string sql, TSqlFragment fragment)
    {
        return CompactSqlWhitespace(
            ExtractFragmentText(
                sql,
                fragment,
                normalizeCoalesce: true,
                applyDisplayAliases: true));
    }

    /// <summary>
    /// SQL断片を帳票向けの大文字表記へ整形
    /// </summary>
    private static string DisplayText(
        string sql,
        TSqlFragment fragment,
        bool uppercaseDateParts = false,
        bool compactUnarySigns = false,
        bool uppercaseOffsetKeywords = false)
    {
        return CompactSqlWhitespace(
            SqlDisplayFormatter.Format(
                sql,
                fragment,
                uppercaseDateParts,
                compactUnarySigns,
                uppercaseOffsetKeywords,
                CurrentDisplayAliases.Value?.Replacements));
    }

    /// <summary>
    /// AST位置から表記を変更せず元SQLを取得
    /// </summary>
    private static string RawFragmentText(string sql, TSqlFragment fragment)
    {
        return ExtractFragmentText(
            sql,
            fragment,
            normalizeCoalesce: false,
            applyDisplayAliases: false);
    }

    /// <summary>
    /// AST位置からSQL断片を取得し描画用キーワードを正規化
    /// </summary>
    private static string ExtractFragmentText(
        string sql,
        TSqlFragment fragment,
        bool normalizeCoalesce,
        bool applyDisplayAliases)
    {
        if (fragment.StartOffset < 0 || fragment.FragmentLength <= 0 ||
            fragment.StartOffset + fragment.FragmentLength > sql.Length)
        {
            return string.Empty;
        }

        var text = sql.Substring(fragment.StartOffset, fragment.FragmentLength);
        if (normalizeCoalesce)
        {
            text = NormalizeCoalesceKeywords(text, fragment.StartOffset, fragment);
        }
        if (applyDisplayAliases && CurrentDisplayAliases.Value is not null)
        {
            text = ApplyTextReplacements(
                text,
                fragment.StartOffset,
                CurrentDisplayAliases.Value.Replacements);
        }

        return text.Trim();
    }

    /// <summary>
    /// 元SQL上の絶対位置を使い、表示用置換を断片の後方から適用
    /// </summary>
    private static string ApplyTextReplacements(
        string text,
        int fragmentStartOffset,
        IReadOnlyList<SqlTextReplacement> replacements)
    {
        var fragmentEndOffset = fragmentStartOffset + text.Length;
        foreach (var replacement in replacements
            .Where(item =>
                item.Offset >= fragmentStartOffset &&
                item.Offset + item.Length <= fragmentEndOffset)
            .OrderByDescending(item => item.Offset))
        {
            var relativeOffset = replacement.Offset - fragmentStartOffset;
            if (relativeOffset < 0 ||
                replacement.Length <= 0 ||
                relativeOffset + replacement.Length > text.Length)
            {
                continue;
            }

            text = text.Remove(relativeOffset, replacement.Length)
                .Insert(relativeOffset, replacement.Value);
        }

        return text;
    }

    /// <summary>
    /// ASTで識別したCOALESCEだけを大文字へ統一
    /// </summary>
    private static string NormalizeCoalesceKeywords(
        string text,
        int fragmentStartOffset,
        TSqlFragment fragment)
    {
        const string keyword = "COALESCE";
        if (!text.Contains(keyword, StringComparison.OrdinalIgnoreCase))
        {
            return text;
        }

        var collector = new CoalesceExpressionCollector();
        fragment.Accept(collector);
        if (collector.StartOffsets.Count == 0)
        {
            return text;
        }

        var characters = text.ToCharArray();
        foreach (var startOffset in collector.StartOffsets)
        {
            var relativeOffset = startOffset - fragmentStartOffset;
            if (relativeOffset < 0 || relativeOffset + keyword.Length > text.Length ||
                string.Compare(
                    text,
                    relativeOffset,
                    keyword,
                    0,
                    keyword.Length,
                    StringComparison.OrdinalIgnoreCase) != 0)
            {
                continue;
            }

            keyword.CopyTo(0, characters, relativeOffset, keyword.Length);
        }

        return new string(characters);
    }

    /// <summary>
    /// 未対応SQLを行単位で出力し原因を末尾へ追加
    /// </summary>
    private static OutputSheetPlan CreateFallback(
        string sql,
        string reason,
        TSqlFragment? causeFragment = null,
        int? causeStartLine = null,
        int? causeEndLine = null)
    {
        var text = sql.Trim('\r', '\n');
        var lines = text.Length == 0
            ? Array.Empty<string>()
            : text
                .Replace("\r\n", "\n", StringComparison.Ordinal)
                .Replace('\r', '\n')
                .Split('\n');
        var cells = lines
            .Select((line, index) => new OutputCell(index + 1, 1, line))
            .ToList();
        var reasonRow = lines.Length == 0 ? 1 : lines.Length + 2;
        var (sourceStartLine, sourceEndLine) = ResolveFallbackSourceLines(
            causeFragment,
            causeStartLine,
            causeEndLine);
        var (queryStartRow, queryEndRow) = ResolveFallbackQueryRows(
            sql,
            lines.Length,
            sourceStartLine,
            sourceEndLine);
        var message = queryStartRow.HasValue && queryEndRow.HasValue
            ? FormatFallbackMessage(reason, queryStartRow.Value, queryEndRow.Value)
            : "フォールバック原因: " + reason;
        cells.Add(new OutputCell(reasonRow, 1, message));
        return new OutputSheetPlan(
            cells,
            [],
            reasonRow,
            true,
            reason,
            queryStartRow,
            queryEndRow,
            FallbackSourceStartLine: sourceStartLine,
            FallbackSourceEndLine: sourceEndLine);
    }

    /// <summary>
    /// 原因断片または明示行を元SQL上の論理行範囲へ変換
    /// </summary>
    private static (int? StartLine, int? EndLine) ResolveFallbackSourceLines(
        TSqlFragment? causeFragment,
        int? causeStartLine,
        int? causeEndLine)
    {
        if (causeStartLine.HasValue)
        {
            return (
                Math.Max(causeStartLine.Value, 1),
                Math.Max((causeEndLine ?? causeStartLine).Value, causeStartLine.Value));
        }

        if (causeFragment is null ||
            causeFragment.LastTokenIndex < 0 ||
            causeFragment.LastTokenIndex >= causeFragment.ScriptTokenStream.Count)
        {
            return (null, null);
        }

        var endLine = causeFragment.ScriptTokenStream[causeFragment.LastTokenIndex].Line;
        return (Math.Max(causeFragment.StartLine, 1), Math.Max(endLine, causeFragment.StartLine));
    }

    /// <summary>
    /// 元SQL上の原因行を先頭改行除去後のフォールバック出力行へ変換
    /// </summary>
    private static (int? StartRow, int? EndRow) ResolveFallbackQueryRows(
        string sql,
        int outputLineCount,
        int? sourceStartLine,
        int? sourceEndLine)
    {
        if (outputLineCount == 0)
        {
            return (null, null);
        }

        if (!sourceStartLine.HasValue)
        {
            return (1, outputLineCount);
        }

        var removedLeadingLines = CountLeadingLineBreaks(sql);
        var startRow = Math.Clamp(
            sourceStartLine.Value - removedLeadingLines,
            1,
            outputLineCount);
        var endRow = Math.Clamp(
            (sourceEndLine ?? sourceStartLine).Value - removedLeadingLines,
            startRow,
            outputLineCount);
        return (startRow, endRow);
    }

    /// <summary>
    /// 出力時に除去する先頭改行の行数を取得
    /// </summary>
    private static int CountLeadingLineBreaks(string sql)
    {
        var count = 0;
        var index = 0;
        while (index < sql.Length)
        {
            if (sql[index] == '\r')
            {
                count++;
                index += index + 1 < sql.Length && sql[index + 1] == '\n' ? 2 : 1;
            }
            else if (sql[index] == '\n')
            {
                count++;
                index++;
            }
            else
            {
                break;
            }
        }

        return count;
    }

    /// <summary>
    /// 原因とアウトプットシート上の対象行を表示
    /// </summary>
    private static string FormatFallbackMessage(string reason, int startRow, int endRow)
    {
        var rowText = startRow == endRow
            ? $"{startRow}行目"
            : $"{startRow}～{endRow}行目";
        return $"フォールバック原因: {reason}（対象クエリ: アウトプット① {rowText}）";
    }

    /// <summary>
    /// parserの構文エラーを利用者向けの原因へ変換
    /// </summary>
    private static string BuildParseErrorReason(IList<ParseError> errors)
    {
        if (errors.Count == 0)
        {
            return "T-SQLを解析できませんでした";
        }

        var error = errors[0];
        return $"T-SQL解析エラー (行{error.Line}, 列{error.Column}): {error.Message}";
    }

    /// <summary>
    /// parserの構文エラー理由と先頭エラー行をフォールバック計画へ設定
    /// </summary>
    private static OutputSheetPlan CreateParseFallback(string sql, IList<ParseError> errors)
    {
        var errorLine = errors.Count > 0 ? errors[0].Line : (int?)null;
        return CreateFallback(
            sql,
            BuildParseErrorReason(errors),
            causeStartLine: errorLine,
            causeEndLine: errorLine);
    }

    /// <summary>
    /// 未対応ステートメントの表示名を取得
    /// </summary>
    private static string StatementKind(TSqlStatement? statement)
    {
        if (statement is null)
        {
            return "なし";
        }

        return statement.GetType().Name.Replace("Statement", string.Empty, StringComparison.Ordinal).ToUpperInvariant();
    }

    /// <summary>
    /// INSERTソースの表示名を取得
    /// </summary>
    private static string InsertSourceKind(TSqlFragment source)
    {
        return source.GetType().Name switch
        {
            "ValuesInsertSource" => "VALUES",
            "ExecuteInsertSource" => "EXECUTE",
            _ => source.GetType().Name
        };
    }

    /// <summary>
    /// 単純な列参照は移送元へ、式は移送方法へ配置して参照列を移送元へ列挙
    /// </summary>
    private static TransferItem CreateTransferItem(
        string sql,
        string target,
        string expressionText,
        ScalarExpression expression)
    {
        var unwrapped = UnwrapScalarExpression(expression);
        if (unwrapped is ColumnReferenceExpression directColumn)
        {
            return new TransferItem(
                target,
                FragmentText(sql, directColumn),
                string.Empty,
                expression,
                DisplayAliases: CurrentDisplayAliases.Value);
        }

        if (unwrapped is ScalarSubquery)
        {
            return new TransferItem(
                target,
                expressionText,
                string.Empty,
                expression,
                DisplayAliases: CurrentDisplayAliases.Value);
        }

        var sources = CollectTransferSources(
            sql,
            ColumnReferenceCollector.Collect(expression));
        return new TransferItem(
            target,
            string.Join("、", sources),
            expressionText,
            expression,
            DisplayAliases: CurrentDisplayAliases.Value);
    }

    /// <summary>
    /// スカラー式を囲む括弧を除去して直接列参照か判定可能にする
    /// </summary>
    private static ScalarExpression UnwrapScalarExpression(ScalarExpression expression)
    {
        while (expression is ParenthesisExpression parenthesized)
        {
            expression = parenthesized.Expression;
        }

        return expression;
    }

    /// <summary>
    /// 更新系のCASEを移送方法へ配置し、戻り値の列参照を移送元へ列挙
    /// </summary>
    private static TransferItem CreateCaseAwareTransferItem(
        string sql,
        string target,
        string expressionText,
        ScalarExpression expression)
    {
        var directCases = DirectCaseExpressions(expression);
        if (directCases.Count > 0)
        {
            var sources = CollectCaseAwareSources(sql, expression);
            return new TransferItem(
                target,
                string.Join("、", sources),
                expressionText,
                expression,
                RenderCaseInMethod: true,
                DirectCases: directCases,
                DisplayAliases: CurrentDisplayAliases.Value);
        }

        return CreateTransferItem(sql, target, expressionText, expression);
    }

    /// <summary>
    /// UPDATE SETの直接スカラーサブクエリをSQ参照へ変換
    /// </summary>
    private static TransferItem CreateUpdateTransferItem(
        string sql,
        AssignmentSetClause clause,
        IReadOnlyList<SubqueryInfo> directSubqueries,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var target = FragmentText(sql, clause.Column);
        if (clause.NewValue is ScalarSubquery scalarSubquery)
        {
            var subquery = directSubqueries.FirstOrDefault(candidate =>
                ReferenceEquals(candidate.QueryExpression, scalarSubquery.QueryExpression));
            if (subquery is not null &&
                TryResolveSingleSelectFieldName(
                    sql,
                    scalarSubquery.QueryExpression,
                    mappings,
                    out var fieldName))
            {
                return new TransferItem(target, $"{subquery.Name}.{fieldName}", string.Empty);
            }
        }

        return CreateCaseAwareTransferItem(
            sql,
            target,
            FragmentText(sql, clause.NewValue),
            clause.NewValue);
    }

    /// <summary>
    /// CASEの条件式を除外し、外側の式とTHEN/ELSEが参照する列を出現順に収集
    /// </summary>
    private static IReadOnlyList<string> CollectCaseAwareSources(
        string sql,
        ScalarExpression expression)
    {
        return CollectTransferSources(
            sql,
            CaseAwareSourceCollector.Collect(expression));
    }

    /// <summary>
    /// 列参照を出現順・重複なしの移送元表示へ変換
    /// </summary>
    private static IReadOnlyList<string> CollectTransferSources(
        string sql,
        IEnumerable<ColumnReferenceExpression> columns)
    {
        var sources = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in columns)
        {
            var source = FragmentText(sql, column);
            if (source == "*" || source.EndsWith(".*", StringComparison.Ordinal))
            {
                continue;
            }
            if (seen.Add(source))
            {
                sources.Add(source);
            }
        }

        return sources;
    }

    /// <summary>
    /// 単一取得項目の出力名を列名または別名から解決
    /// </summary>
    private static bool TryResolveSingleSelectFieldName(
        string sql,
        QueryExpression expression,
        IReadOnlyList<MappingDefinition> mappings,
        out string fieldName)
    {
        fieldName = string.Empty;
        if (UnwrapQueryExpression(expression) is not QuerySpecification query ||
            query.SelectElements.Count != 1 ||
            query.SelectElements[0] is not SelectScalarExpression scalar)
        {
            return false;
        }

        string fieldId;
        if (scalar.ColumnName is not null)
        {
            fieldId = FragmentText(sql, scalar.ColumnName);
        }
        else if (scalar.Expression is ColumnReferenceExpression column &&
            column.MultiPartIdentifier?.Identifiers.Count > 0)
        {
            fieldId = column.MultiPartIdentifier.Identifiers[^1].Value;
        }
        else
        {
            return false;
        }

        fieldName = ResolveOutputFieldName(string.Empty, fieldId, mappings);
        return true;
    }

    /// <summary>
    /// 式内のCASEを収集し、サブクエリ内部は別フレームへ委ねる
    /// </summary>
    private sealed class CaseExpressionCollector : TSqlFragmentVisitor
    {
        private readonly List<ScalarExpression> _items = [];

        /// <summary>
        /// SQL断片に含まれるCASEを取得
        /// </summary>
        public static IReadOnlyList<ScalarExpression> Collect(TSqlFragment fragment)
        {
            var collector = new CaseExpressionCollector();
            fragment.Accept(collector);
            return collector._items;
        }

        /// <summary>
        /// 検索CASEを追加
        /// </summary>
        public override void ExplicitVisit(SearchedCaseExpression node)
        {
            _items.Add(node);
            base.ExplicitVisit(node);
        }

        /// <summary>
        /// 単純CASEを追加
        /// </summary>
        public override void ExplicitVisit(SimpleCaseExpression node)
        {
            _items.Add(node);
            base.ExplicitVisit(node);
        }

        /// <summary>
        /// サブクエリ内のCASEを親式から除外
        /// </summary>
        public override void ExplicitVisit(ScalarSubquery node)
        {
        }
    }

    /// <summary>
    /// サブクエリを内側から収集し、出力名を割り当てる
    /// </summary>
    private sealed class SubqueryCollector : TSqlFragmentVisitor
    {
        private readonly List<SubqueryInfo> _items = [];
        private readonly HashSet<(int StartOffset, int Length)> _seen = [];

        /// <summary>
        /// SQL断片から出力対象サブクエリを収集
        /// </summary>
        public static IReadOnlyList<SubqueryInfo> Collect(TSqlFragment fragment)
        {
            var collector = new SubqueryCollector();
            fragment.Accept(collector);
            return collector._items;
        }

        /// <summary>
        /// CTEを名前付きサブクエリとして追加
        /// </summary>
        public override void ExplicitVisit(CommonTableExpression node)
        {
            Add(node.QueryExpression, node.ExpressionName.Value, true);
        }

        /// <summary>
        /// 派生テーブルを無名サブクエリとして追加
        /// </summary>
        public override void ExplicitVisit(QueryDerivedTable node)
        {
            Add(node.QueryExpression, node.Alias?.Value, false);
            base.ExplicitVisit(node);
        }

        /// <summary>
        /// スカラーサブクエリを追加
        /// </summary>
        public override void ExplicitVisit(ScalarSubquery node)
        {
            Add(node.QueryExpression, null, false);
            base.ExplicitVisit(node);
        }

        /// <summary>
        /// EXISTS内のサブクエリを追加
        /// </summary>
        public override void ExplicitVisit(ExistsPredicate node)
        {
            Add(node.Subquery.QueryExpression, null, false);
            base.ExplicitVisit(node);
        }

        /// <summary>
        /// IN内のサブクエリを追加
        /// </summary>
        public override void ExplicitVisit(InPredicate node)
        {
            if (node.Subquery is not null)
            {
                Add(node.Subquery.QueryExpression, null, false);
            }

            base.ExplicitVisit(node);
        }

        /// <summary>
        /// 子を先に収集してから重複なく追加
        /// </summary>
        private void Add(QueryExpression query, string? explicitName, bool isNamed)
        {
            var key = (query.StartOffset, query.FragmentLength);
            if (_seen.Contains(key))
            {
                return;
            }

            query.Accept(this);
            if (!_seen.Add(key))
            {
                return;
            }

            var name = explicitName ?? $"SQ{_items.Count + 1}";
            _items.Add(new SubqueryInfo(query, name, isNamed));
        }
    }

    private sealed record SubqueryInfo(QueryExpression QueryExpression, string Name, bool IsNamed);

    private sealed record SubqueryReplacement(
        string Name,
        IReadOnlyList<string> SourceTexts,
        Regex ParenthesizedNamePattern);

    /// <summary>
    /// 未修飾列の候補を、列IDとテーブル和名の索引から解決
    /// </summary>
    private sealed class ColumnQualificationIndex
    {
        private readonly Dictionary<string, List<MappingDefinition>> mappingsByField =
            new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, HashSet<string>> anchoredTableIdsByName =
            new(StringComparer.OrdinalIgnoreCase);

        public ColumnQualificationIndex(IReadOnlyList<MappingDefinition> mappings)
        {
            foreach (var mapping in mappings)
            {
                AddFieldMappings(mapping);
                AddTableNameAnchor(mapping);
            }
        }

        public bool IsEmpty => mappingsByField.Count == 0;

        public bool AssociatesWithTable(string fieldId, NamedTableReference table)
        {
            if (!mappingsByField.TryGetValue(fieldId, out var candidates))
            {
                return false;
            }

            return candidates.Any(mapping =>
                MappingBelongsToTable(mapping, table) ||
                StandaloneMappingBelongsToTable(mapping, table));
        }

        private void AddFieldMappings(MappingDefinition mapping)
        {
            var fieldKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            AddFieldKey(fieldKeys, mapping.FieldId);
            AddFieldKey(fieldKeys, mapping.FieldName);
            AddFieldKey(fieldKeys, mapping.ParserFieldId);
            foreach (var fieldKey in fieldKeys)
            {
                if (!mappingsByField.TryGetValue(fieldKey, out var fieldMappings))
                {
                    fieldMappings = [];
                    mappingsByField.Add(fieldKey, fieldMappings);
                }
                fieldMappings.Add(mapping);
            }
        }

        private static void AddFieldKey(ISet<string> fieldKeys, string fieldKey)
        {
            if (fieldKey.Length > 0)
            {
                fieldKeys.Add(fieldKey);
            }
        }

        private void AddTableNameAnchor(MappingDefinition mapping)
        {
            if (mapping.TableId == "-" || !IsUsableTableName(mapping.TableName))
            {
                return;
            }

            var tableName = mapping.TableName.Trim();
            if (!anchoredTableIdsByName.TryGetValue(tableName, out var tableIds))
            {
                tableIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                anchoredTableIdsByName.Add(tableName, tableIds);
            }
            tableIds.Add(mapping.TableId);
        }

        private bool StandaloneMappingBelongsToTable(
            MappingDefinition mapping,
            NamedTableReference table)
        {
            if (mapping.TableId != "-" || !IsUsableTableName(mapping.TableName) ||
                !anchoredTableIdsByName.TryGetValue(mapping.TableName.Trim(), out var tableIds))
            {
                return false;
            }

            return tableIds.Contains(table.Alias?.Value ?? string.Empty) ||
                tableIds.Contains(table.SchemaObject.BaseIdentifier.Value);
        }

        private static bool IsUsableTableName(string tableName)
        {
            var value = tableName.Trim();
            return value.Length > 0 &&
                value != "-" &&
                !value.Contains(MissingName, StringComparison.Ordinal);
        }
    }

    private sealed record QualificationInsertion(
        string Prefix,
        OutputReplacementQualification Qualification);

    private sealed record QualificationResult(
        string Sql,
        IReadOnlyList<OutputReplacementQualification> Replacements);

    /// <summary>
    /// SQL断片に含まれるSELECT本体を列挙
    /// </summary>
    private sealed class QuerySpecificationCollector : TSqlFragmentVisitor
    {
        private readonly List<QuerySpecification> _queries = [];

        /// <summary>
        /// SQL断片内のSELECT本体を取得
        /// </summary>
        public static IReadOnlyList<QuerySpecification> Collect(TSqlFragment fragment)
        {
            var collector = new QuerySpecificationCollector();
            fragment.Accept(collector);
            return collector._queries;
        }

        /// <summary>
        /// SELECT本体を追加して、ネストしたSELECTも引き続き探索
        /// </summary>
        public override void ExplicitVisit(QuerySpecification node)
        {
            _queries.Add(node);
            base.ExplicitVisit(node);
        }
    }

    /// <summary>
    /// parser専用フィールドIDを末尾識別子に持つ列参照を収集
    /// </summary>
    private sealed class ParserFieldColumnCollector : TSqlFragmentVisitor
    {
        private readonly ISet<string> parserFieldIds;
        private readonly List<ParserFieldColumn> columns = [];

        private ParserFieldColumnCollector(IEnumerable<string> parserFieldIds)
        {
            this.parserFieldIds = parserFieldIds.ToHashSet(StringComparer.Ordinal);
        }

        public static IReadOnlyList<ParserFieldColumn> Collect(
            TSqlFragment fragment,
            IEnumerable<string> parserFieldIds)
        {
            var collector = new ParserFieldColumnCollector(parserFieldIds);
            fragment.Accept(collector);
            return collector.columns;
        }

        public override void ExplicitVisit(ColumnReferenceExpression node)
        {
            var identifiers = node.MultiPartIdentifier?.Identifiers;
            if (identifiers is { Count: > 0 } &&
                parserFieldIds.Contains(identifiers[^1].Value))
            {
                columns.Add(new ParserFieldColumn(node, identifiers[^1]));
            }
            base.ExplicitVisit(node);
        }
    }

    /// <summary>
    /// 式内の列参照を出現順に収集し、サブクエリ内部は別フレームへ委ねる
    /// </summary>
    private sealed class ColumnReferenceCollector : TSqlFragmentVisitor
    {
        private readonly List<ColumnReferenceExpression> _columns = [];

        /// <summary>
        /// SQL断片に含まれる列参照を取得
        /// </summary>
        public static IReadOnlyList<ColumnReferenceExpression> Collect(TSqlFragment expression)
        {
            var collector = new ColumnReferenceCollector();
            expression.Accept(collector);
            return collector._columns;
        }

        /// <summary>
        /// 列参照を追加
        /// </summary>
        public override void ExplicitVisit(ColumnReferenceExpression node)
        {
            _columns.Add(node);
        }

        /// <summary>
        /// スカラーサブクエリ内部の列は親の移送元へ含めない
        /// </summary>
        public override void ExplicitVisit(ScalarSubquery node)
        {
        }
    }

    /// <summary>
    /// CASE条件を除き、外側の式とCASE戻り値に含まれる列参照を出現順に収集
    /// </summary>
    private sealed class CaseAwareSourceCollector : TSqlFragmentVisitor
    {
        private readonly List<ColumnReferenceExpression> _columns = [];

        /// <summary>
        /// CASEを含む式から移送元となる列参照を収集
        /// </summary>
        public static IReadOnlyList<ColumnReferenceExpression> Collect(ScalarExpression expression)
        {
            var collector = new CaseAwareSourceCollector();
            expression.Accept(collector);
            return collector._columns;
        }

        /// <summary>
        /// 列参照を収集
        /// </summary>
        public override void ExplicitVisit(ColumnReferenceExpression node)
        {
            _columns.Add(node);
        }

        /// <summary>
        /// 検索CASEはWHEN条件を除き、THENとELSEだけを収集
        /// </summary>
        public override void ExplicitVisit(SearchedCaseExpression node)
        {
            foreach (var clause in node.WhenClauses)
            {
                clause.ThenExpression.Accept(this);
            }
            node.ElseExpression?.Accept(this);
        }

        /// <summary>
        /// 単純CASEは入力式とWHEN値を除き、THENとELSEだけを収集
        /// </summary>
        public override void ExplicitVisit(SimpleCaseExpression node)
        {
            foreach (var clause in node.WhenClauses)
            {
                clause.ThenExpression.Accept(this);
            }
            node.ElseExpression?.Accept(this);
        }
    }

    /// <summary>
    /// 子サブクエリを除いた列修飾子を出現順に収集
    /// </summary>
    private sealed class DirectColumnQualifierCollector : TSqlFragmentVisitor
    {
        private readonly List<string> _identifiers = [];
        private readonly HashSet<string> _seen = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// SQL断片直下の列修飾子を収集
        /// </summary>
        public static IReadOnlyList<string> Collect(TSqlFragment fragment)
        {
            var collector = new DirectColumnQualifierCollector();
            fragment.Accept(collector);
            return collector._identifiers;
        }

        /// <summary>
        /// 2要素以上の列参照からテーブル修飾子を追加
        /// </summary>
        public override void ExplicitVisit(ColumnReferenceExpression node)
        {
            var identifiers = node.MultiPartIdentifier?.Identifiers;
            if (identifiers is null || identifiers.Count < 2)
            {
                return;
            }

            var identifier = identifiers[^2].Value;
            if (_seen.Add(identifier))
            {
                _identifiers.Add(identifier);
            }
        }

        /// <summary>
        /// スカラーサブクエリ内部を収集対象から除外
        /// </summary>
        public override void ExplicitVisit(ScalarSubquery node)
        {
        }

        /// <summary>
        /// EXISTSサブクエリ内部を収集対象から除外
        /// </summary>
        public override void ExplicitVisit(ExistsPredicate node)
        {
        }

        /// <summary>
        /// INサブクエリ内部を収集対象から除外
        /// </summary>
        public override void ExplicitVisit(InPredicate node)
        {
            node.Expression.Accept(this);
            foreach (var value in node.Values)
            {
                value.Accept(this);
            }
        }

        /// <summary>
        /// FROM内の派生クエリを収集対象から除外
        /// </summary>
        public override void ExplicitVisit(QueryDerivedTable node)
        {
        }
    }

    /// <summary>
    /// SQL断片内のCOALESCE開始位置を収集
    /// </summary>
    private sealed class CoalesceExpressionCollector : TSqlFragmentVisitor
    {
        public List<int> StartOffsets { get; } = [];

        /// <summary>
        /// COALESCEの開始位置を追加
        /// </summary>
        public override void ExplicitVisit(CoalesceExpression node)
        {
            StartOffsets.Add(node.StartOffset);
            base.ExplicitVisit(node);
        }
    }

    /// <summary>
    /// 複合クエリの各分岐へ、フレーム内で一意な表示用テーブル別名を割り当て
    /// </summary>
    private sealed class BinaryDisplayAliasPlan
    {
        private readonly IReadOnlyList<BranchDisplayAliases> branches;

        private BinaryDisplayAliasPlan(IReadOnlyList<BranchDisplayAliases> branches)
        {
            this.branches = branches;
        }

        public static BinaryDisplayAliasPlan Create(
            IReadOnlyList<QuerySpecification> queryBranches)
        {
            var reservedAliases = UsedTableIdentifierCollector.Collect(queryBranches);
            var firstPhysicalByAlias = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
            var displayByBinding = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
            var assignments = new List<BranchAliasAssignment>(queryBranches.Count);
            var collidingAliases = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var branch in queryBranches)
            {
                var tableBindings = new Dictionary<int, DisplayAliasBinding>();
                var renamedAliases = new Dictionary<string, string>(
                    StringComparer.OrdinalIgnoreCase);
                var declarationReplacements = new List<SqlTextReplacement>();
                var namedTables = branch.FromClause?.TableReferences
                    .SelectMany(EnumerateNamedTables)
                    .ToArray() ?? [];
                var duplicateAlias = branch.FromClause?.TableReferences
                    .SelectMany(EnumerateTableIdentifiers)
                    .Where(identifier => identifier.Length > 0)
                    .GroupBy(identifier => identifier, StringComparer.OrdinalIgnoreCase)
                    .FirstOrDefault(group => group.Count() > 1);
                if (duplicateAlias is not null)
                {
                    throw new UnsupportedOutputException(
                        "同一分岐内でテーブル別名が重複しています: " + duplicateAlias.Key,
                        branch);
                }
                foreach (var table in namedTables)
                {
                    if (table.Alias is null)
                    {
                        continue;
                    }

                    var sourceAlias = table.Alias.Value;
                    var physicalId = PhysicalTableIdentity(table);
                    var bindingKey = sourceAlias + "\u001f" + physicalId;
                    if (!displayByBinding.TryGetValue(bindingKey, out var displayAlias))
                    {
                        if (!firstPhysicalByAlias.TryGetValue(sourceAlias, out var firstPhysicalId))
                        {
                            firstPhysicalByAlias.Add(sourceAlias, physicalId);
                            displayAlias = sourceAlias;
                        }
                        else if (string.Equals(
                            firstPhysicalId,
                            physicalId,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            displayAlias = sourceAlias;
                        }
                        else
                        {
                            displayAlias = AllocateDisplayAlias(reservedAliases);
                        }

                        displayByBinding.Add(bindingKey, displayAlias);
                    }

                    tableBindings[table.StartOffset] = new DisplayAliasBinding(
                        sourceAlias,
                        displayAlias,
                        table.SchemaObject.BaseIdentifier.Value,
                        PreservePhysicalTableId: false);
                    if (!string.Equals(
                        sourceAlias,
                        displayAlias,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        renamedAliases[sourceAlias] = displayAlias;
                        collidingAliases.Add(sourceAlias);
                        declarationReplacements.Add(new SqlTextReplacement(
                            table.Alias.StartOffset,
                            table.Alias.FragmentLength,
                            DisplayAliasReferenceCollector.DisplayIdentifier(
                                displayAlias,
                                table.Alias.QuoteType)));
                    }
                }

                assignments.Add(new BranchAliasAssignment(
                    branch,
                    tableBindings,
                    renamedAliases,
                    declarationReplacements));
            }

            var branchAliases = assignments
                .Select(assignment => new BranchDisplayAliases(
                    assignment.Query,
                    new DisplayAliasContext(
                        assignment.TableBindings.ToDictionary(
                            item => item.Key,
                            item => item.Value with
                            {
                                PreservePhysicalTableId = collidingAliases.Contains(
                                    item.Value.SourceAlias)
                            }),
                        assignment.DeclarationReplacements
                            .Concat(DisplayAliasReferenceCollector.Collect(
                                assignment.Query,
                                assignment.RenamedAliases))
                            .GroupBy(item => item.Offset)
                            .Select(group => group.Last())
                            .OrderBy(item => item.Offset)
                            .ToArray())))
                .ToArray();
            return new BinaryDisplayAliasPlan(branchAliases);
        }

        public IReadOnlyList<SqlTextReplacement> Replacements => branches
            .SelectMany(branch => branch.Aliases.Replacements)
            .GroupBy(item => item.Offset)
            .Select(group => group.Last())
            .OrderBy(item => item.Offset)
            .ToArray();

        public DisplayAliasContext? ContextFor(QuerySpecification query)
        {
            return branches
                .FirstOrDefault(item => ReferenceEquals(item.Query, query))
                ?.Aliases;
        }

        public DisplayAliasContext? ContextContaining(TSqlFragment fragment)
        {
            return branches
                .Where(item => ContainsFragment(item.Query, fragment))
                .OrderBy(item => item.Query.FragmentLength)
                .FirstOrDefault()
                ?.Aliases;
        }

        private static string AllocateDisplayAlias(ISet<string> reservedAliases)
        {
            for (var number = 1; ; number++)
            {
                var candidate = $"tb{number}";
                if (reservedAliases.Add(candidate))
                {
                    return candidate;
                }
            }
        }

        private static string PhysicalTableIdentity(NamedTableReference table)
        {
            return string.Join(
                ".",
                table.SchemaObject.Identifiers.Select(identifier => identifier.Value));
        }
    }

    /// <summary>
    /// 対応済みステートメント内の複合クエリからSQL全体用の別名置換を収集
    /// </summary>
    private sealed class BinaryDisplayAliasReplacementCollector : TSqlFragmentVisitor
    {
        private readonly HashSet<(int Offset, int Length)> visited = [];
        private readonly List<SqlTextReplacement> replacements = [];

        public static IReadOnlyList<SqlTextReplacement> Collect(
            IEnumerable<TSqlStatement> statements)
        {
            var collector = new BinaryDisplayAliasReplacementCollector();
            foreach (var statement in statements)
            {
                statement.Accept(collector);
            }

            return collector.replacements
                .GroupBy(item => item.Offset)
                .Select(group => group.Last())
                .OrderBy(item => item.Offset)
                .ToArray();
        }

        public override void ExplicitVisit(BinaryQueryExpression node)
        {
            if (!visited.Add((node.StartOffset, node.FragmentLength)))
            {
                return;
            }

            var branches = new List<QuerySpecification>();
            var separators = new List<string>();
            AddBinaryBranches(node, branches, separators);
            if (branches.Count == 0)
            {
                return;
            }

            replacements.AddRange(BinaryDisplayAliasPlan.Create(branches).Replacements);

            // 同じ集合演算ツリーは一度だけ割り当て、分岐内の独立した集合演算だけを探索する。
            foreach (var branch in branches)
            {
                branch.Accept(this);
            }
            node.OrderByClause?.Accept(this);
            node.OffsetClause?.Accept(this);
        }
    }

    /// <summary>
    /// 1分岐内で元SQL位置と表示用別名を対応付け
    /// </summary>
    private sealed class DisplayAliasContext
    {
        private readonly IReadOnlyDictionary<int, DisplayAliasBinding> tableBindings;
        private readonly IReadOnlyDictionary<string, DisplayAliasBinding> bindingsBySourceAlias;

        public DisplayAliasContext(
            IReadOnlyDictionary<int, DisplayAliasBinding> tableBindings,
            IReadOnlyList<SqlTextReplacement> replacements)
            : this(
                tableBindings,
                tableBindings.Values
                    .GroupBy(binding => binding.SourceAlias, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(
                        group => group.Key,
                        group => group.Last(),
                        StringComparer.OrdinalIgnoreCase),
                replacements)
        {
        }

        private DisplayAliasContext(
            IReadOnlyDictionary<int, DisplayAliasBinding> tableBindings,
            IReadOnlyDictionary<string, DisplayAliasBinding> bindingsBySourceAlias,
            IReadOnlyList<SqlTextReplacement> replacements)
        {
            this.tableBindings = tableBindings;
            this.bindingsBySourceAlias = bindingsBySourceAlias;
            Replacements = replacements;
        }

        public IReadOnlyList<SqlTextReplacement> Replacements { get; }

        public string DisplayAliasFor(NamedTableReference table)
        {
            return tableBindings.TryGetValue(table.StartOffset, out var binding)
                ? binding.DisplayAlias
                : table.Alias?.Value ?? table.SchemaObject.BaseIdentifier.Value;
        }

        public string DisplayAliasFor(string tableId)
        {
            return bindingsBySourceAlias.TryGetValue(tableId, out var binding)
                ? binding.DisplayAlias
                : tableId;
        }

        public string? PhysicalTableIdFor(string tableId)
        {
            return bindingsBySourceAlias.TryGetValue(tableId, out var binding)
                ? binding.PhysicalTableId
                : null;
        }

        public bool ShouldPreservePhysicalTableId(string tableId)
        {
            return bindingsBySourceAlias.TryGetValue(tableId, out var binding) &&
                binding.PreservePhysicalTableId;
        }

        public bool ShouldPreservePhysicalTableId(NamedTableReference table)
        {
            return tableBindings.TryGetValue(table.StartOffset, out var binding) &&
                binding.PreservePhysicalTableId;
        }

        public static DisplayAliasContext? Combine(
            DisplayAliasContext? outer,
            DisplayAliasContext? inner)
        {
            if (outer is null)
            {
                return inner;
            }
            if (inner is null)
            {
                return outer;
            }

            var combinedTableBindings = new Dictionary<int, DisplayAliasBinding>(
                outer.tableBindings);
            foreach (var item in inner.tableBindings)
            {
                combinedTableBindings[item.Key] = item.Value;
            }

            var combinedBindingsBySourceAlias = new Dictionary<string, DisplayAliasBinding>(
                outer.bindingsBySourceAlias,
                StringComparer.OrdinalIgnoreCase);
            foreach (var item in inner.bindingsBySourceAlias)
            {
                combinedBindingsBySourceAlias[item.Key] = item.Value;
            }

            var combinedReplacements = outer.Replacements
                .Concat(inner.Replacements)
                .GroupBy(item => item.Offset)
                .Select(group => group.Last())
                .OrderBy(item => item.Offset)
                .ToArray();
            return new DisplayAliasContext(
                combinedTableBindings,
                combinedBindingsBySourceAlias,
                combinedReplacements);
        }
    }

    /// <summary>
    /// 分岐のローカルスコープを尊重して列修飾子の置換位置を収集
    /// </summary>
    private sealed class DisplayAliasReferenceCollector : TSqlFragmentVisitor
    {
        private readonly QuerySpecification root;
        private readonly List<SqlTextReplacement> replacements = [];
        private IReadOnlyDictionary<string, string> activeAliases;

        private DisplayAliasReferenceCollector(
            QuerySpecification root,
            IReadOnlyDictionary<string, string> renamedAliases)
        {
            this.root = root;
            activeAliases = renamedAliases;
        }

        public static IReadOnlyList<SqlTextReplacement> Collect(
            QuerySpecification query,
            IReadOnlyDictionary<string, string> renamedAliases)
        {
            if (renamedAliases.Count == 0)
            {
                return [];
            }

            var collector = new DisplayAliasReferenceCollector(query, renamedAliases);
            query.Accept(collector);
            return collector.replacements
                .GroupBy(item => item.Offset)
                .Select(group => group.Last())
                .OrderBy(item => item.Offset)
                .ToArray();
        }

        public override void ExplicitVisit(QuerySpecification node)
        {
            var previousAliases = activeAliases;
            if (!ReferenceEquals(node, root))
            {
                var localAliases = (node.FromClause?.TableReferences
                    .SelectMany(EnumerateTableIdentifiers) ?? [])
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
                activeAliases = previousAliases
                    .Where(item => !localAliases.Contains(item.Key))
                    .ToDictionary(
                        item => item.Key,
                        item => item.Value,
                        StringComparer.OrdinalIgnoreCase);
            }

            base.ExplicitVisit(node);
            activeAliases = previousAliases;
        }

        public override void ExplicitVisit(ColumnReferenceExpression node)
        {
            var identifiers = node.MultiPartIdentifier?.Identifiers;
            if (identifiers is { Count: >= 2 })
            {
                AddReplacement(identifiers[^2]);
            }
            base.ExplicitVisit(node);
        }

        public override void ExplicitVisit(SelectStarExpression node)
        {
            var identifiers = node.Qualifier?.Identifiers;
            if (identifiers is { Count: > 0 })
            {
                AddReplacement(identifiers[^1]);
            }
            base.ExplicitVisit(node);
        }

        private void AddReplacement(Identifier identifier)
        {
            if (activeAliases.TryGetValue(identifier.Value, out var displayAlias))
            {
                replacements.Add(new SqlTextReplacement(
                    identifier.StartOffset,
                    identifier.FragmentLength,
                    DisplayIdentifier(displayAlias, identifier.QuoteType)));
            }
        }

        public static string DisplayIdentifier(string value, QuoteType quoteType)
        {
            return quoteType switch
            {
                QuoteType.SquareBracket => $"[{value.Replace("]", "]]", StringComparison.Ordinal)}]",
                QuoteType.DoubleQuote => $"\"{value.Replace("\"", "\"\"", StringComparison.Ordinal)}\"",
                _ => value
            };
        }
    }

    /// <summary>
    /// 複合クエリ内で使用済みの実表名・別名・派生表別名を収集
    /// </summary>
    private sealed class UsedTableIdentifierCollector : TSqlFragmentVisitor
    {
        private readonly HashSet<string> identifiers = new(
            StringComparer.OrdinalIgnoreCase);

        public static HashSet<string> Collect(
            IEnumerable<QuerySpecification> queryBranches)
        {
            var collector = new UsedTableIdentifierCollector();
            foreach (var branch in queryBranches)
            {
                branch.Accept(collector);
            }
            return collector.identifiers;
        }

        public override void ExplicitVisit(NamedTableReference node)
        {
            identifiers.Add(node.Alias?.Value ?? node.SchemaObject.BaseIdentifier.Value);
            base.ExplicitVisit(node);
        }

        public override void ExplicitVisit(QueryDerivedTable node)
        {
            if (node.Alias is not null)
            {
                identifiers.Add(node.Alias.Value);
            }
            base.ExplicitVisit(node);
        }

        public override void ExplicitVisit(InlineDerivedTable node)
        {
            if (node.Alias is not null)
            {
                identifiers.Add(node.Alias.Value);
            }
            base.ExplicitVisit(node);
        }
    }

    private sealed class DisplayAliasScope(DisplayAliasContext? previous) : IDisposable
    {
        private bool disposed;

        public void Dispose()
        {
            if (disposed)
            {
                return;
            }

            CurrentDisplayAliases.Value = previous;
            disposed = true;
        }
    }

    private sealed record BranchAliasAssignment(
        QuerySpecification Query,
        IReadOnlyDictionary<int, DisplayAliasBinding> TableBindings,
        IReadOnlyDictionary<string, string> RenamedAliases,
        IReadOnlyList<SqlTextReplacement> DeclarationReplacements);

    private sealed record DisplayAliasBinding(
        string SourceAlias,
        string DisplayAlias,
        string PhysicalTableId,
        bool PreservePhysicalTableId);

    private sealed record MissingTableDisplayCandidate(
        string SourceValue,
        string PhysicalTableId,
        string ReplacementSuffix);

    private sealed record BranchDisplayAliases(
        QuerySpecification Query,
        DisplayAliasContext Aliases);

    private sealed record ParserFieldOccurrence(
        int Offset,
        MappingDefinition Mapping);

    private sealed record ParserFieldColumn(
        ColumnReferenceExpression Column,
        Identifier FieldIdentifier);

    private sealed record SqlTextPosition(int Line, int Column);

    private sealed class UnsupportedOutputException(
        string message,
        TSqlFragment? fragment = null) : Exception(message)
    {
        public TSqlFragment? Fragment { get; } = fragment;
    }

    private sealed record TransferItem(
        string Target,
        string Source,
        string Method,
        ScalarExpression? Expression = null,
        bool RenderCaseInMethod = false,
        IReadOnlyList<ScalarExpression>? DirectCases = null,
        DisplayAliasContext? DisplayAliases = null);

    private sealed record ConditionPart(string Connector, BooleanExpression Expression);

    private sealed record JoinTableDisplay(string Identifier, string Display);

    private sealed record CaseConditionPart(
        string Connector,
        int ConnectorDepth,
        BooleanExpression Expression,
        int ExpressionDepth,
        int OpeningParentheses = 0,
        int ClosingParentheses = 0);

    private sealed record ConditionLayout(
        int ConnectorColumn,
        int FirstValueColumn,
        int ConnectedValueColumn,
        int ConnectedGroupOpenColumn);
}

/// <summary>
/// 変換定義シートから渡す和名定義
/// </summary>
public sealed record MappingDefinition(
    string TableId,
    string TableName,
    string FieldId,
    string FieldName,
    string ParserFieldId = "");
