using Microsoft.SqlServer.TransactSql.ScriptDom;

namespace SqlAnalysisFormatter.Parser;

internal sealed record PhysicalTableUsage(
    IReadOnlyList<string> InputTableIds,
    IReadOnlyList<string> OutputTableIds);

internal static class PhysicalTableUsageCollector
{
    public static PhysicalTableUsage Collect(IEnumerable<TSqlStatement> statements)
    {
        ArgumentNullException.ThrowIfNull(statements);

        var inputs = new OrderedTableIds();
        var outputs = new OrderedTableIds();
        foreach (var statement in statements)
        {
            CollectStatement(statement, inputs, outputs);
        }

        return new PhysicalTableUsage(inputs.Items, outputs.Items);
    }

    private static void CollectStatement(
        TSqlStatement statement,
        OrderedTableIds inputs,
        OrderedTableIds outputs)
    {
        if (statement is not SelectStatement and
            not InsertStatement and
            not UpdateStatement and
            not DeleteStatement)
        {
            return;
        }

        var cteDefinitions = CollectCteDefinitions(statement);
        var namedTables = NamedTableCollector.Collect(statement);
        var specification = GetDataModificationSpecification(statement);
        TableReference? target = statement switch
        {
            InsertStatement insert => insert.InsertSpecification.Target,
            UpdateStatement update => update.UpdateSpecification.Target,
            DeleteStatement delete => delete.DeleteSpecification.Target,
            _ => null
        };
        var outputIntoTarget = specification?.OutputIntoClause?.IntoTable;
        var targetBinding = FindTargetBinding(statement, target);
        var resolvedTarget = ResolveTarget(statement, target, cteDefinitions);

        // UPDATE/DELETEで更新対象の列をSET/WHERE/ON/OUTPUTなどから読む場合は、
        // 同じ物理表を入力と出力の両方へ載せる。
        if (statement is UpdateStatement or DeleteStatement &&
            ReadsModificationTarget(statement, target, targetBinding, resolvedTarget))
        {
            inputs.Add(resolvedTarget);
        }

        // 更新対象そのものと、UPDATE/DELETEのFROM句で対象別名を束縛する出現は重複収集しない。
        // 自己結合やサブクエリなど、同じ物理表の独立した出現は除外せず入力へ残す。
        var suppressInputs = statement is InsertStatement valuesInsertStatement &&
            valuesInsertStatement.InsertSpecification.InsertSource is ValuesInsertSource;
        if (!suppressInputs)
        {
            foreach (var table in namedTables
                .Where(table => !ReferenceEquals(table, target))
                .Where(table => !ReferenceEquals(table, targetBinding))
                .Where(table => !ReferenceEquals(table, outputIntoTarget))
                .Where(table => !IsCteReference(table, cteDefinitions)))
            {
                inputs.Add(PhysicalId(table));
            }
        }

        switch (statement)
        {
            case SelectStatement select when select.Into is not null:
                outputs.Add(select.Into.BaseIdentifier.Value);
                break;
            case InsertStatement:
            case UpdateStatement:
            case DeleteStatement:
                outputs.Add(resolvedTarget);
                break;
        }
        if (outputIntoTarget is not null)
        {
            outputs.Add(ResolveTableReference(outputIntoTarget, cteDefinitions));
        }
    }

    private static bool ReadsModificationTarget(
        TSqlStatement statement,
        TableReference? target,
        TableReference? targetBinding,
        string resolvedTarget)
    {
        var targetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        AddTargetNames(targetNames, target);
        AddTargetNames(targetNames, targetBinding);
        if (resolvedTarget.Length > 0)
        {
            targetNames.Add(resolvedTarget);
        }

        var assignmentTargets = statement is UpdateStatement update
            ? update.UpdateSpecification.SetClauses
                .OfType<AssignmentSetClause>()
                .Where(clause => clause.AssignmentKind == AssignmentKind.Equals)
                .Where(clause => clause.Column is not null)
                .Select(clause => clause.Column)
                .ToHashSet()
            : [];
        var queryScopes = QueryScopeCollector.Collect(statement);

        foreach (var column in ColumnReferenceCollector.Collect(statement))
        {
            if (assignmentTargets.Contains(column))
            {
                continue;
            }

            var identifiers = column.MultiPartIdentifier?.Identifiers;
            if (identifiers is not { Count: > 0 })
            {
                continue;
            }

            if (identifiers.Count > 1)
            {
                var qualifier = identifiers[^2].Value;
                var isTargetQualifier = targetNames.Contains(qualifier) ||
                    qualifier.Equals("inserted", StringComparison.OrdinalIgnoreCase) ||
                    qualifier.Equals("deleted", StringComparison.OrdinalIgnoreCase);
                if (isTargetQualifier &&
                    !queryScopes.Any(scope => scope.Contains(column.StartOffset) &&
                        scope.TableQualifiers.Contains(qualifier)) &&
                    !IsBoundToOtherModificationSource(statement, qualifier, targetBinding))
                {
                    return true;
                }
                continue;
            }

            // 無修飾列は、FROMがないか対象テーブルだけの場合に限り更新対象列として扱う。
            if (!queryScopes.Any(scope => scope.Contains(column.StartOffset)) &&
                HasOnlyTargetSource(statement, targetBinding))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsBoundToOtherModificationSource(
        TSqlStatement statement,
        string qualifier,
        TableReference? targetBinding)
    {
        var sources = GetModificationFromClause(statement)?.TableReferences
            .SelectMany(EnumerateSourceTables)
            ?? [];
        var foundSource = false;
        foreach (var source in sources)
        {
            var sourceQualifier = source is TableReferenceWithAlias { Alias: not null } aliased
                ? aliased.Alias.Value
                : source is NamedTableReference named ? PhysicalId(named) : string.Empty;
            if (!sourceQualifier.Equals(qualifier, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            foundSource = true;
            if (ReferenceEquals(source, targetBinding))
            {
                return false;
            }
        }

        return foundSource;
    }

    private static bool HasOnlyTargetSource(TSqlStatement statement, TableReference? targetBinding)
    {
        var fromClause = GetModificationFromClause(statement);
        if (fromClause is null)
        {
            return true;
        }

        // ponytail: without column metadata, multi-source unqualified names stay ambiguous; add catalog-aware resolution if exact ownership is needed.
        var sources = fromClause.TableReferences
            .SelectMany(EnumerateSourceTables)
            .Take(2)
            .ToArray();
        return sources.Length == 1 && ReferenceEquals(sources[0], targetBinding);
    }

    private static IEnumerable<TableReference> EnumerateSourceTables(TableReference table)
    {
        switch (table)
        {
            case JoinTableReference join:
                foreach (var source in EnumerateSourceTables(join.FirstTableReference))
                {
                    yield return source;
                }
                foreach (var source in EnumerateSourceTables(join.SecondTableReference))
                {
                    yield return source;
                }
                break;
            case JoinParenthesisTableReference parenthesized:
                foreach (var source in EnumerateSourceTables(parenthesized.Join))
                {
                    yield return source;
                }
                break;
            default:
                yield return table;
                break;
        }
    }

    private static void AddTargetNames(ISet<string> targetNames, TableReference? tableReference)
    {
        if (tableReference is NamedTableReference named)
        {
            targetNames.Add(PhysicalId(named));
        }
        if (tableReference is TableReferenceWithAlias aliased && aliased.Alias is not null)
        {
            targetNames.Add(aliased.Alias.Value);
        }
    }

    private static DataModificationSpecification? GetDataModificationSpecification(
        TSqlStatement statement)
    {
        return statement switch
        {
            InsertStatement insert => insert.InsertSpecification,
            UpdateStatement update => update.UpdateSpecification,
            DeleteStatement delete => delete.DeleteSpecification,
            _ => null
        };
    }

    private static TableReference? FindTargetBinding(
        TSqlStatement statement,
        TableReference? target)
    {
        return FindTargetBinding(target, GetModificationFromClause(statement));
    }

    internal static TableReference? FindTargetBinding(
        TableReference? target,
        FromClause? fromClause)
    {
        if (target is not NamedTableReference namedTarget)
        {
            return null;
        }

        var targetId = PhysicalId(namedTarget);
        if (namedTarget.SchemaObject.Identifiers.Count == 1)
        {
            var aliasMatch = FindTableReferenceByAlias(fromClause?.TableReferences, targetId);
            if (aliasMatch is not null)
            {
                return aliasMatch;
            }
        }

        var physicalMatches = FindTableReferencesByPhysicalId(fromClause?.TableReferences, targetId)
            .Take(2)
            .ToArray();
        return physicalMatches.Length == 1 ? physicalMatches[0] : null;
    }

    private static IEnumerable<NamedTableReference> FindTableReferencesByPhysicalId(
        IEnumerable<TableReference>? tableReferences,
        string physicalId)
    {
        if (tableReferences is null)
        {
            yield break;
        }

        foreach (var tableReference in tableReferences)
        {
            foreach (var match in FindTableReferencesByPhysicalId(tableReference, physicalId))
            {
                yield return match;
            }
        }
    }

    private static IEnumerable<NamedTableReference> FindTableReferencesByPhysicalId(
        TableReference tableReference,
        string physicalId)
    {
        switch (tableReference)
        {
            case NamedTableReference named when string.Equals(
                PhysicalId(named), physicalId, StringComparison.OrdinalIgnoreCase):
                yield return named;
                break;
            case JoinTableReference join:
                foreach (var match in FindTableReferencesByPhysicalId(join.FirstTableReference, physicalId))
                {
                    yield return match;
                }
                foreach (var match in FindTableReferencesByPhysicalId(join.SecondTableReference, physicalId))
                {
                    yield return match;
                }
                break;
            case JoinParenthesisTableReference parenthesized:
                foreach (var match in FindTableReferencesByPhysicalId(parenthesized.Join, physicalId))
                {
                    yield return match;
                }
                break;
        }
    }

    private static FromClause? GetModificationFromClause(TSqlStatement statement)
    {
        return statement switch
        {
            UpdateStatement update => update.UpdateSpecification.FromClause,
            DeleteStatement delete => delete.DeleteSpecification.FromClause,
            _ => null
        };
    }

    private static IReadOnlyDictionary<string, CommonTableExpression> CollectCteDefinitions(
        TSqlStatement statement)
    {
        var visitor = new CommonTableExpressionCollector();
        statement.Accept(visitor);
        return visitor.Definitions;
    }

    private static bool IsCteReference(
        NamedTableReference table,
        IReadOnlyDictionary<string, CommonTableExpression> cteDefinitions)
    {
        return table.SchemaObject.Identifiers.Count == 1 &&
            cteDefinitions.ContainsKey(table.SchemaObject.BaseIdentifier.Value);
    }

    private static string PhysicalId(NamedTableReference table)
    {
        return table.SchemaObject.BaseIdentifier?.Value?.Trim() ?? string.Empty;
    }

    private static string ResolveTarget(
        TSqlStatement statement,
        TableReference? target,
        IReadOnlyDictionary<string, CommonTableExpression> cteDefinitions)
    {
        if (target is not NamedTableReference namedTarget)
        {
            return string.Empty;
        }

        var targetId = PhysicalId(namedTarget);
        if (targetId.Length == 0)
        {
            return string.Empty;
        }

        if (namedTarget.SchemaObject.Identifiers.Count > 1)
        {
            return targetId;
        }

        if (cteDefinitions.TryGetValue(targetId, out var cte))
        {
            return ResolveUniquePhysicalTable(cte.QueryExpression, cteDefinitions);
        }

        var aliasMatch = FindTargetBinding(
            namedTarget,
            GetModificationFromClause(statement));
        if (aliasMatch is not null)
        {
            return ResolveTableReference(aliasMatch, cteDefinitions);
        }

        return targetId;
    }

    private static TableReference? FindTableReferenceByAlias(
        IEnumerable<TableReference>? tableReferences,
        string alias)
    {
        if (tableReferences is null)
        {
            return null;
        }

        foreach (var tableReference in tableReferences)
        {
            var match = FindTableReferenceByAlias(tableReference, alias);
            if (match is not null)
            {
                return match;
            }
        }

        return null;
    }

    private static TableReference? FindTableReferenceByAlias(
        TableReference tableReference,
        string alias)
    {
        if (tableReference is TableReferenceWithAlias aliased &&
            aliased.Alias is not null &&
            string.Equals(aliased.Alias.Value, alias, StringComparison.OrdinalIgnoreCase))
        {
            return tableReference;
        }

        return tableReference switch
        {
            JoinTableReference join =>
                FindTableReferenceByAlias(join.FirstTableReference, alias) ??
                FindTableReferenceByAlias(join.SecondTableReference, alias),
            JoinParenthesisTableReference parenthesized =>
                FindTableReferenceByAlias(parenthesized.Join, alias),
            _ => null
        };
    }

    private static string ResolveTableReference(
        TableReference tableReference,
        IReadOnlyDictionary<string, CommonTableExpression> cteDefinitions)
    {
        return tableReference switch
        {
            NamedTableReference named when IsCteReference(named, cteDefinitions) =>
                ResolveUniquePhysicalTable(
                    cteDefinitions[PhysicalId(named)].QueryExpression,
                    cteDefinitions),
            NamedTableReference named => PhysicalId(named),
            QueryDerivedTable derived =>
                ResolveUniquePhysicalTable(derived.QueryExpression, cteDefinitions),
            JoinParenthesisTableReference parenthesized =>
                ResolveTableReference(parenthesized.Join, cteDefinitions),
            JoinTableReference join => ResolveUniquePhysicalTable(join, cteDefinitions),
            _ => string.Empty
        };
    }

    private static string ResolveUniquePhysicalTable(
        TSqlFragment fragment,
        IReadOnlyDictionary<string, CommonTableExpression> cteDefinitions)
    {
        var physicalTables = new OrderedTableIds();
        AddUpdatableRowsetTables(
            fragment,
            cteDefinitions,
            new HashSet<string>(StringComparer.OrdinalIgnoreCase),
            physicalTables);
        return physicalTables.Items.Count == 1 ? physicalTables.Items[0] : string.Empty;
    }

    private static void AddUpdatableRowsetTables(
        TSqlFragment fragment,
        IReadOnlyDictionary<string, CommonTableExpression> cteDefinitions,
        ISet<string> visitedCtes,
        OrderedTableIds physicalTables)
    {
        switch (fragment)
        {
            case QueryParenthesisExpression parenthesizedQuery:
                AddUpdatableRowsetTables(
                    parenthesizedQuery.QueryExpression,
                    cteDefinitions,
                    visitedCtes,
                    physicalTables);
                break;
            case QuerySpecification query when query.FromClause is not null:
                foreach (var table in query.FromClause.TableReferences)
                {
                    AddUpdatableRowsetTables(
                        table,
                        cteDefinitions,
                        visitedCtes,
                        physicalTables);
                }
                break;
            case NamedTableReference cteReference when IsCteReference(cteReference, cteDefinitions):
                var cteId = PhysicalId(cteReference);
                if (visitedCtes.Add(cteId))
                {
                    AddUpdatableRowsetTables(
                        cteDefinitions[cteId].QueryExpression,
                        cteDefinitions,
                        visitedCtes,
                        physicalTables);
                }
                break;
            case NamedTableReference namedTable:
                var tableId = PhysicalId(namedTable);
                physicalTables.Add(tableId);
                break;
            case QueryDerivedTable derived:
                AddUpdatableRowsetTables(
                    derived.QueryExpression,
                    cteDefinitions,
                    visitedCtes,
                    physicalTables);
                break;
            case JoinParenthesisTableReference parenthesizedJoin:
                AddUpdatableRowsetTables(
                    parenthesizedJoin.Join,
                    cteDefinitions,
                    visitedCtes,
                    physicalTables);
                break;
            case JoinTableReference join:
                AddUpdatableRowsetTables(
                    join.FirstTableReference,
                    cteDefinitions,
                    visitedCtes,
                    physicalTables);
                AddUpdatableRowsetTables(
                    join.SecondTableReference,
                    cteDefinitions,
                    visitedCtes,
                    physicalTables);
                break;
        }
    }

    private sealed class CommonTableExpressionCollector : TSqlFragmentVisitor
    {
        public Dictionary<string, CommonTableExpression> Definitions { get; } =
            new(StringComparer.OrdinalIgnoreCase);

        public override void ExplicitVisit(CommonTableExpression node)
        {
            Definitions.TryAdd(node.ExpressionName.Value, node);
            base.ExplicitVisit(node);
        }
    }

    private sealed class NamedTableCollector : TSqlFragmentVisitor
    {
        private readonly List<NamedTableReference> tables = [];

        public static IReadOnlyList<NamedTableReference> Collect(TSqlFragment fragment)
        {
            var visitor = new NamedTableCollector();
            fragment.Accept(visitor);
            return visitor.tables
                .OrderBy(table => table.StartOffset)
                .ToArray();
        }

        public override void ExplicitVisit(NamedTableReference node)
        {
            tables.Add(node);
            base.ExplicitVisit(node);
        }
    }

    private sealed class ColumnReferenceCollector : TSqlFragmentVisitor
    {
        private readonly List<ColumnReferenceExpression> columns = [];

        public static IReadOnlyList<ColumnReferenceExpression> Collect(TSqlFragment fragment)
        {
            var visitor = new ColumnReferenceCollector();
            fragment.Accept(visitor);
            return visitor.columns;
        }

        public override void ExplicitVisit(ColumnReferenceExpression node)
        {
            columns.Add(node);
            base.ExplicitVisit(node);
        }
    }

    private sealed class QueryScopeCollector : TSqlFragmentVisitor
    {
        private readonly List<QueryScope> scopes = [];

        public static IReadOnlyList<QueryScope> Collect(TSqlFragment fragment)
        {
            var visitor = new QueryScopeCollector();
            fragment.Accept(visitor);
            return visitor.scopes;
        }

        public override void ExplicitVisit(QuerySpecification node)
        {
            var qualifiers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (node.FromClause is not null)
            {
                foreach (var table in node.FromClause.TableReferences.SelectMany(EnumerateSourceTables))
                {
                    if (table is TableReferenceWithAlias { Alias: not null } aliased)
                    {
                        qualifiers.Add(aliased.Alias.Value);
                    }
                    else if (table is NamedTableReference named)
                    {
                        qualifiers.Add(PhysicalId(named));
                    }
                }
            }

            scopes.Add(new QueryScope(
                node.StartOffset,
                node.StartOffset + node.FragmentLength,
                qualifiers));
            base.ExplicitVisit(node);
        }
    }

    private sealed record QueryScope(int StartOffset, int EndOffset, HashSet<string> TableQualifiers)
    {
        public bool Contains(int offset) => offset >= StartOffset && offset < EndOffset;
    }

    private sealed class OrderedTableIds
    {
        private readonly HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<string> items = [];

        public IReadOnlyList<string> Items => items;

        public void Add(string? tableId)
        {
            var normalized = tableId?.Trim() ?? string.Empty;
            if (normalized.Length > 0 && seen.Add(normalized))
            {
                items.Add(normalized);
            }
        }
    }
}
