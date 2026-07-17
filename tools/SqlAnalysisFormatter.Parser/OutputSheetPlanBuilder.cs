using Microsoft.SqlServer.TransactSql.ScriptDom;
using System.Text;
using System.Text.RegularExpressions;

namespace SqlAnalysisFormatter.Parser;

/// <summary>
/// T-SQL ASTをアウトプットシートの描画計画へ変換
/// </summary>
public static class OutputSheetPlanBuilder
{
    private const string MissingName = "(和名未取得)";

    /// <summary>
    /// 和名変換済みSQLから描画計画を作成
    /// </summary>
    public static OutputSheetPlan Build(string sql, IReadOnlyList<MappingDefinition> mappings)
    {
        ArgumentNullException.ThrowIfNull(sql);
        ArgumentNullException.ThrowIfNull(mappings);

        var plan = BuildCore(sql, mappings);
        return ParserFieldIdentifierRestorer.Restore(plan, mappings);
    }

    /// <summary>
    /// SQLを解析して復元前の描画計画を作成
    /// </summary>
    private static OutputSheetPlan BuildCore(string sql, IReadOnlyList<MappingDefinition> mappings)
    {
        var parser = new TSql160Parser(initialQuotedIdentifiers: false);
        using var reader = new StringReader(sql);
        var fragment = parser.Parse(reader, out var errors);
        if (errors.Count > 0 || fragment is not TSqlScript script)
        {
            return CreateFallback(sql, BuildParseErrorReason(errors));
        }

        var statement = script.Batches.FirstOrDefault()?.Statements.FirstOrDefault();
        try
        {
            return statement switch
            {
                SelectStatement selectStatement => BuildSelectStatement(sql, selectStatement, mappings),
                InsertStatement insertStatement => BuildInsert(sql, insertStatement, mappings),
                UpdateStatement updateStatement => BuildUpdate(sql, updateStatement, mappings),
                DeleteStatement deleteStatement => BuildDelete(sql, deleteStatement, mappings),
                _ => CreateFallback(sql, "未対応のステートメント: " + StatementKind(statement))
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

        var (subqueries, plans) = BuildLeadingSubqueryPlans(sql, selectStatement, mappings);

        var wholeChildren = DirectChildSubqueries(selectStatement.QueryExpression, subqueries);
        var wholePlan = BuildQueryExpression(
            sql,
            selectStatement.QueryExpression,
            mappings,
            "＜DB入出力項目定義＞",
            wholeChildren.Where(child => !child.IsNamed).Select(child => child.Name));
        plans.Add(ReplaceSubqueries(wholePlan, sql, wholeChildren));

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
        var transfers = BuildSelectTransfers(sql, targetColumns, sourceQuery.SelectElements);
        var transferPlan = BuildDataTransferPlan(
            sql,
            targetName,
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
            IReadOnlyList<MappingDefinition> mappings)
    {
        var subqueries = SubqueryCollector.Collect(fragment);
        var plans = new List<OutputSheetPlan>(subqueries.Count + 1);
        foreach (var subquery in subqueries)
        {
            var children = DirectChildSubqueries(subquery.QueryExpression, subqueries);
            var plan = BuildQueryExpression(
                sql,
                subquery.QueryExpression,
                mappings,
                $"サブクエリ[{subquery.Name}]",
                children.Where(child => !child.IsNamed).Select(child => child.Name));
            plans.Add(ReplaceSubqueries(plan, sql, children));
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
    /// SELECT INTOの出力別名を変換定義から和名へ解決
    /// </summary>
    private static string ResolveOutputFieldName(
        string targetId,
        string fieldId,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var mapping = mappings
            .Where(item => string.Equals(item.FieldId, fieldId, StringComparison.OrdinalIgnoreCase))
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
                    string.Empty));
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
        if (sourceExpression is not QuerySpecification sourceQuery)
        {
            return CreateFallback(sql, "未対応のINSERT形式: SELECTの集合演算", selectSource.Select);
        }

        if (specification.Columns.Count != sourceQuery.SelectElements.Count)
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

        var targets = specification.Columns
            .Select(column => FragmentText(sql, column))
            .ToArray();
        var transfers = BuildSelectTransfers(sql, targets, sourceQuery.SelectElements);
        var targetDisplay = BuildTargetTableDisplay(
            specification.Target,
            mappings,
            includeIdentifier: false);
        var transferPlan = BuildDataTransferPlan(
            sql,
            targetDisplay,
            transfers,
            null,
            null,
            mappings);
        plans.Add(ReplaceSubqueries(transferPlan, sql, sourceChildren));

        return CombinePlans(plans);
    }

    /// <summary>
    /// 単一行のINSERT VALUESをデータ移送表へ変換
    /// </summary>
    private static OutputSheetPlan BuildInsertValues(
        string sql,
        InsertSpecification specification,
        ValuesInsertSource valuesSource,
        IReadOnlyList<MappingDefinition> mappings)
    {
        if (valuesSource.IsDefaultValues)
        {
            return CreateFallback(sql, "未対応のINSERT形式: DEFAULT VALUES", valuesSource);
        }

        if (valuesSource.RowValues.Count != 1)
        {
            return CreateFallback(sql, "INSERT VALUESの複数行入力は未対応", valuesSource);
        }

        var values = valuesSource.RowValues[0].ColumnValues;
        if (specification.Columns.Count == 0 || specification.Columns.Count != values.Count)
        {
            return CreateFallback(
                sql,
                "INSERT VALUESの対象列数と値数が一致しません",
                valuesSource);
        }

        var transfers = new List<TransferItem>(values.Count);
        for (var index = 0; index < values.Count; index++)
        {
            transfers.Add(CreateCaseAwareTransferItem(
                sql,
                FragmentText(sql, specification.Columns[index]),
                FragmentText(sql, values[index]),
                values[index]));
        }

        var targetDisplay = BuildTargetTableDisplay(
            specification.Target,
            mappings,
            includeIdentifier: false);
        return BuildDataTransferPlan(
            sql,
            targetDisplay,
            transfers,
            null,
            null,
            mappings);
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
            var startRow = row;
            var transferGroups = new List<OutputSection>();
            cells.Add(new OutputCell(row, 1, "項目"));
            cells.Add(new OutputCell(row, 19, "移送元"));
            cells.Add(new OutputCell(row, 37, "移送方法ほか"));
            row++;
            foreach (var transfer in transfers)
            {
                var itemStartRow = row;
                cells.Add(new OutputCell(row, 1, transfer.Target));
                if (transfer.Expression is not null &&
                    DirectCaseExpressions(transfer.Expression).Count > 0)
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
                            detailColumn: 52)
                        : WriteScalarExpression(
                            cells,
                            sql,
                            transfer.Expression,
                            row,
                            displayName: null,
                            valueColumn: 19,
                            markerColumn: 35,
                            detailColumn: 37);
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
            sections.AddRange(transferGroups);
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
    /// 更新対象、FROM句、サブクエリを更新系の参照テーブル一覧へ統合
    /// </summary>
    private static string BuildModificationTableList(
        TableReference target,
        FromClause? fromClause,
        IReadOnlyList<MappingDefinition> mappings,
        IEnumerable<string> additionalTables)
    {
        var targetDisplay = BuildTargetTableDisplay(target, mappings, includeIdentifier: true);
        var displays = new[] { targetDisplay }
            .Concat(BuildTableDisplays(fromClause, mappings, additionalTables))
            .Distinct(StringComparer.OrdinalIgnoreCase);
        return string.Join("、", displays);
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

        var tableId = named.Alias?.Value ?? named.SchemaObject.BaseIdentifier.Value;
        var tableName = ResolveTableName(named, mappings);
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
        IEnumerable<string> additionalTables)
    {
        expression = UnwrapQueryExpression(expression);
        return expression switch
        {
            QuerySpecification query => BuildSelect(sql, query, mappings, title, additionalTables),
            BinaryQueryExpression binary => BuildBinaryQuery(sql, binary, mappings, title, additionalTables),
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
        IEnumerable<string> additionalTables)
    {
        var branches = new List<QuerySpecification>();
        var separators = new List<string>();
        AddBinaryBranches(binary, branches, separators);
        if (branches.Count == 0)
        {
            return CreateFallback(RawFragmentText(sql, binary), "複合クエリの分岐を取得できませんでした");
        }

        var tableDisplays = branches
            .SelectMany(branch => BuildTableDisplays(branch.FromClause, mappings, []))
            .Concat(additionalTables)
            .Distinct(StringComparer.OrdinalIgnoreCase);
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
            var branchPlan = BuildSelect(sql, branches[index], mappings, title, []);
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
        IReadOnlyList<SubqueryInfo> subqueries)
    {
        if (subqueries.Count == 0)
        {
            return plan;
        }

        var cells = plan.Cells.Select(cell =>
        {
            var value = cell.Value;
            foreach (var subquery in subqueries)
            {
                var sourceTexts = new[]
                {
                    RawFragmentText(sql, subquery.QueryExpression),
                    FragmentText(sql, subquery.QueryExpression),
                    DisplayText(sql, subquery.QueryExpression)
                };
                foreach (var sourceText in sourceTexts
                    .Where(text => text.Length > 0)
                    .Distinct(StringComparer.Ordinal))
                {
                    value = value.Replace(sourceText, subquery.Name, StringComparison.Ordinal);
                }
                value = Regex.Replace(
                    value,
                    @"\(\s*" + Regex.Escape(subquery.Name) + @"\s*\)",
                    "(" + subquery.Name + ")",
                    RegexOptions.CultureInvariant);
            }
            return cell with { Value = value };
        }).ToArray();
        return plan with { Cells = cells };
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
            fallbackQueryEndRow);
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

        if (query.OffsetClause is not null)
        {
            var startRow = row;
            cells.Add(new OutputCell(row, 1, "取得範囲"));
            var cases = DirectCaseExpressions(query.OffsetClause);
            if (cases.Count == 0)
            {
                cells.Add(new OutputCell(row, 7, FragmentText(sql, query.OffsetClause)));
                row++;
            }
            else
            {
                cells.Add(new OutputCell(
                    row,
                    7,
                    RenderExpressionWithCasePlaceholders(sql, query.OffsetClause, cases)));
                cells.Add(new OutputCell(row, 15, "※"));
                row += WriteEmbeddedCaseBranches(cells, sql, cases, row, 17);
            }
            sections.Add(new OutputSection(OutputSectionKind.Standard, startRow, row - 1));
        }

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

        if (query.OrderByClause is not null && query.OrderByClause.OrderByElements.Count > 0)
        {
            var startRow = row;
            for (var index = 0; index < query.OrderByClause.OrderByElements.Count; index++)
            {
                var element = query.OrderByClause.OrderByElements[index];
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

        return new OutputSheetPlan(
            cells,
            sections,
            row - 1,
            false);
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
        var columns = ColumnReferenceCollector.Collect(scalar.Expression);
        if (columns.Count == 0)
        {
            return alias;
        }

        var namedTables = query.FromClause?.TableReferences
            .SelectMany(EnumerateNamedTables)
            .ToArray() ?? [];
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
        int detailColumn = 32)
    {
        var cases = DirectCaseExpressions(expression);
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
        return WriteEmbeddedCaseBranches(cells, sql, cases, row, detailColumn + 2);
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
        return parts;
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
            row += WriteCaseBranches(cells, sql, cases[index], row, column + 2);
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
        IReadOnlyList<ScalarExpression> cases)
    {
        var value = DisplayText(sql, expression);
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
        sections.Add(new OutputSection(OutputSectionKind.Standard, startRow, row - 1));
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

            var leftTables = EnumerateJoinDisplays(join.FirstTableReference, mappings).ToArray();
            var rightTables = EnumerateJoinDisplays(join.SecondTableReference, mappings).ToArray();
            if (leftTables.Length == 0 || rightTables.Length == 0)
            {
                var unsupportedTable = leftTables.Length == 0
                    ? join.FirstTableReference
                    : join.SecondTableReference;
                throw new UnsupportedOutputException(
                    "JOIN対象のテーブル形式は未対応: " + unsupportedTable.GetType().Name,
                    unsupportedTable);
            }

            var leftTable = leftTables[^1];
            var rightTable = rightTables[0];
            var joinText = $"＜{leftTable} {JoinTypeText(join.QualifiedJoinType)} {rightTable}＞";
            cells.Add(new OutputCell(row, 17, joinText));
            row++;

            var parts = FlattenBooleanExpression(join.SearchCondition);
            for (var conditionIndex = 0; conditionIndex < parts.Count; conditionIndex++)
            {
                if (conditionIndex > 0)
                {
                    cells.Add(new OutputCell(row, 7, parts[conditionIndex].Connector));
                }

                if (TryWriteConditionCases(
                    cells,
                    sql,
                    parts[conditionIndex].Expression,
                    17,
                    row,
                    out var consumedRows))
                {
                    row += consumedRows;
                }
                else
                {
                    cells.Add(new OutputCell(row, 17, DisplayText(sql, parts[conditionIndex].Expression)));
                    row++;
                }
            }
        }

        sections.Add(new OutputSection(OutputSectionKind.Standard, startRow, row - 1));
    }

    /// <summary>
    /// 連鎖JOINを内側から列挙
    /// </summary>
    private static IEnumerable<QualifiedJoin> EnumerateJoins(TableReference table)
    {
        if (table is not QualifiedJoin join)
        {
            yield break;
        }

        foreach (var innerJoin in EnumerateJoins(join.FirstTableReference))
        {
            yield return innerJoin;
        }

        yield return join;
    }

    /// <summary>
    /// JOIN片側の実テーブルと派生テーブル名を左から列挙
    /// </summary>
    private static IEnumerable<string> EnumerateJoinDisplays(
        TableReference table,
        IReadOnlyList<MappingDefinition> mappings)
    {
        switch (table)
        {
            case NamedTableReference named:
                yield return BuildTableDisplay(named, mappings);
                break;
            case QueryDerivedTable queryDerived:
                yield return queryDerived.Alias?.Value ?? MissingName;
                break;
            case InlineDerivedTable inlineDerived:
                yield return $"派生テーブル[{inlineDerived.Alias?.Value ?? MissingName}]";
                break;
            case QualifiedJoin join:
                foreach (var display in EnumerateJoinDisplays(join.FirstTableReference, mappings))
                {
                    yield return display;
                }
                foreach (var display in EnumerateJoinDisplays(join.SecondTableReference, mappings))
                {
                    yield return display;
                }
                break;
        }
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
        return displays.Length == 0 ? "なし" : string.Join("、", displays);
    }

    /// <summary>
    /// FROM句と追加参照名から参照テーブル一覧を作成
    /// </summary>
    private static string BuildTableList(
        FromClause? fromClause,
        IReadOnlyList<MappingDefinition> mappings,
        IEnumerable<string> additionalTables)
    {
        var displays = BuildTableDisplays(fromClause, mappings, additionalTables);
        return displays.Count == 0 ? "なし" : string.Join("、", displays);
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
            case QualifiedJoin join:
                foreach (var display in EnumerateTableDisplays(
                    join.FirstTableReference,
                    mappings,
                    allowStandaloneTableName))
                {
                    yield return display;
                }
                foreach (var display in EnumerateTableDisplays(
                    join.SecondTableReference,
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
        var tableId = table.Alias?.Value ?? table.SchemaObject.BaseIdentifier.Value;
        var tableName = ResolveTableName(table, mappings);
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

        return $"{tableName}[{tableId}]";
    }

    /// <summary>
    /// テーブル識別子から和名付き表示を作成
    /// </summary>
    private static string BuildTableDisplay(
        string tableId,
        IReadOnlyList<MappingDefinition> mappings)
    {
        return $"{ResolveTableName(tableId, mappings)}[{tableId}]";
    }

    /// <summary>
    /// 一時テーブルは別名で未解決の場合に物理名でも和名を検索
    /// </summary>
    private static string ResolveTableName(
        NamedTableReference table,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var baseTableId = table.SchemaObject.BaseIdentifier.Value;
        var displayTableId = table.Alias?.Value ?? baseTableId;
        var tableName = ResolveTableName(displayTableId, mappings);
        if (tableName == MissingName &&
            baseTableId.StartsWith('#') &&
            !string.Equals(baseTableId, displayTableId, StringComparison.OrdinalIgnoreCase))
        {
            tableName = ResolveTableName(baseTableId, mappings);
        }

        return tableName;
    }

    /// <summary>
    /// テーブルIDから和名を解決
    /// </summary>
    private static string ResolveTableName(
        string tableId,
        IReadOnlyList<MappingDefinition> mappings)
    {
        var tableName = mappings
            .FirstOrDefault(mapping => string.Equals(mapping.TableId, tableId, StringComparison.OrdinalIgnoreCase))
            ?.TableName;
        if (string.IsNullOrWhiteSpace(tableName))
        {
            tableName = MissingName;
        }

        return tableName;
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
            case QualifiedJoin join:
                foreach (var item in EnumerateNamedTables(join.FirstTableReference))
                {
                    yield return item;
                }
                foreach (var item in EnumerateNamedTables(join.SecondTableReference))
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
            case QualifiedJoin join:
                foreach (var identifier in EnumerateTableIdentifiers(join.FirstTableReference))
                {
                    yield return identifier;
                }
                foreach (var identifier in EnumerateTableIdentifiers(join.SecondTableReference))
                {
                    yield return identifier;
                }
                break;
        }
    }

    /// <summary>
    /// 取得項目から式本体を表示
    /// </summary>
    private static string RenderSelectElement(string sql, SelectElement element)
    {
        return element switch
        {
            SelectScalarExpression scalar => FragmentText(sql, scalar.Expression),
            SelectStarExpression star => RenderSelectStar(FragmentText(sql, star)),
            _ => FragmentText(sql, element)
        };
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
            ExtractFragmentText(sql, fragment, normalizeCoalesce: true));
    }

    /// <summary>
    /// SQL断片を帳票向けの大文字表記へ整形
    /// </summary>
    private static string DisplayText(
        string sql,
        TSqlFragment fragment,
        bool uppercaseDateParts = false,
        bool compactUnarySigns = false)
    {
        return CompactSqlWhitespace(
            SqlDisplayFormatter.Format(sql, fragment, uppercaseDateParts, compactUnarySigns));
    }

    /// <summary>
    /// AST位置から表記を変更せず元SQLを取得
    /// </summary>
    private static string RawFragmentText(string sql, TSqlFragment fragment)
    {
        return ExtractFragmentText(sql, fragment, normalizeCoalesce: false);
    }

    /// <summary>
    /// AST位置からSQL断片を取得し描画用キーワードを正規化
    /// </summary>
    private static string ExtractFragmentText(
        string sql,
        TSqlFragment fragment,
        bool normalizeCoalesce)
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

        return text.Trim();
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
        TSqlFragment? causeFragment = null)
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
        var (queryStartRow, queryEndRow) = ResolveFallbackQueryRows(
            sql,
            lines.Length,
            causeFragment);
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
            queryEndRow);
    }

    /// <summary>
    /// 原因断片をフォールバック出力上の行範囲へ変換
    /// </summary>
    private static (int? StartRow, int? EndRow) ResolveFallbackQueryRows(
        string sql,
        int outputLineCount,
        TSqlFragment? causeFragment)
    {
        if (outputLineCount == 0)
        {
            return (null, null);
        }

        if (causeFragment is null ||
            causeFragment.LastTokenIndex < 0 ||
            causeFragment.LastTokenIndex >= causeFragment.ScriptTokenStream.Count)
        {
            return (1, outputLineCount);
        }

        var removedLeadingLines = CountLeadingLineBreaks(sql);
        var startRow = Math.Clamp(causeFragment.StartLine - removedLeadingLines, 1, outputLineCount);
        var endLine = causeFragment.ScriptTokenStream[causeFragment.LastTokenIndex].Line;
        var endRow = Math.Clamp(endLine - removedLeadingLines, startRow, outputLineCount);
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
        return $"フォールバック原因: {reason}（対象クエリ: アウトプットシート {rowText}）";
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
                expression);
        }

        if (unwrapped is ScalarSubquery)
        {
            return new TransferItem(target, expressionText, string.Empty, expression);
        }

        var sources = CollectTransferSources(
            sql,
            ColumnReferenceCollector.Collect(expression));
        return new TransferItem(
            target,
            string.Join("、", sources),
            expressionText,
            expression);
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
        if (DirectCaseExpressions(expression).Count > 0)
        {
            var sources = CollectCaseAwareSources(sql, expression);
            return new TransferItem(
                target,
                string.Join("、", sources),
                expressionText,
                expression,
                RenderCaseInMethod: true);
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
        /// クエリ直下の列修飾子を収集
        /// </summary>
        public static IReadOnlyList<string> Collect(QuerySpecification query)
        {
            var collector = new DirectColumnQualifierCollector();
            query.Accept(collector);
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
        bool RenderCaseInMethod = false);

    private sealed record ConditionPart(string Connector, BooleanExpression Expression);

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
