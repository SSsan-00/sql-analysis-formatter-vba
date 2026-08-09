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

        // 更新対象そのものと、UPDATE/DELETEのFROM句で対象別名を束縛する出現は出力専用とする。
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

        var resolvedTarget = ResolveTarget(statement, target, cteDefinitions);
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
        if (target is not NamedTableReference namedTarget ||
            namedTarget.SchemaObject.Identifiers.Count != 1)
        {
            return null;
        }

        return FindTableReferenceByAlias(
            GetModificationFromClause(statement)?.TableReferences,
            PhysicalId(namedTarget));
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

        var aliasMatch = FindTableReferenceByAlias(
            GetModificationFromClause(statement)?.TableReferences,
            targetId);
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
