using Microsoft.SqlServer.TransactSql.ScriptDom;

namespace SqlAnalysisFormatter.Parser;

/// <summary>
/// ScriptDomを使ったT-SQL AST解析
/// </summary>
public static class TsqlAstParser
{
    /// <summary>
    /// SQLを解析し、クエリ種別とアウトプット用ブロックを返す
    /// </summary>
    public static ParseResult Parse(string sql)
    {
        ArgumentNullException.ThrowIfNull(sql);

        var parser = new TSql160Parser(initialQuotedIdentifiers: false);
        using var reader = new StringReader(sql);
        var fragment = parser.Parse(reader, out var parseErrors);
        if (parseErrors.Count > 0)
        {
            return new ParseResult(
                false,
                "UNKNOWN",
                [CreateWholeBlock(sql)],
                parseErrors.Select(FormatParseError).ToArray());
        }

        var queryType = DetectQueryType(fragment);
        var blocks = QueryBlockCollector.Collect(sql, fragment);
        blocks.Add(CreateWholeBlock(sql));

        return new ParseResult(true, queryType, blocks, []);
    }

    /// <summary>
    /// SQL全体をフォールバック用ブロックへ変換
    /// </summary>
    private static QueryBlock CreateWholeBlock(string sql)
    {
        return new QueryBlock("WHOLE", TrimSqlWhitespace(sql), 0, sql.Length);
    }

    /// <summary>
    /// ScriptDomエラーを行列付き文字列へ変換
    /// </summary>
    private static string FormatParseError(ParseError error)
    {
        return $"{error.Line}:{error.Column} {error.Message}";
    }

    /// <summary>
    /// 先頭文のCRUD種別を判定
    /// </summary>
    private static string DetectQueryType(TSqlFragment fragment)
    {
        if (fragment is not TSqlScript script)
        {
            return "UNKNOWN";
        }

        var statement = script.Batches.FirstOrDefault()?.Statements.FirstOrDefault();
        return statement switch
        {
            SelectStatement => "SELECT",
            InsertStatement => "INSERT",
            UpdateStatement => "UPDATE",
            DeleteStatement => "DELETE",
            ExecuteStatement => "EXEC",
            _ => "UNKNOWN"
        };
    }

    /// <summary>
    /// AST位置から元SQLの文字列を取得
    /// </summary>
    private static string FragmentText(string sql, TSqlFragment fragment)
    {
        if (fragment.StartOffset < 0 || fragment.FragmentLength <= 0)
        {
            return string.Empty;
        }

        if (fragment.StartOffset + fragment.FragmentLength > sql.Length)
        {
            return string.Empty;
        }

        return TrimSqlWhitespace(sql.Substring(fragment.StartOffset, fragment.FragmentLength));
    }

    /// <summary>
    /// SQL断片の前後にある空白文字だけを除去
    /// </summary>
    private static string TrimSqlWhitespace(string source)
    {
        var start = 0;
        while (start < source.Length && IsSqlWhitespace(source[start]))
        {
            start++;
        }

        var end = source.Length - 1;
        while (end >= start && IsSqlWhitespace(source[end]))
        {
            end--;
        }

        return end >= start ? source.Substring(start, end - start + 1) : string.Empty;
    }

    /// <summary>
    /// SQL断片の除去対象空白か判定
    /// </summary>
    private static bool IsSqlWhitespace(char value)
    {
        return value is ' ' or '\t' or '\r' or '\n';
    }

    /// <summary>
    /// AST上のサブクエリを内側から収集
    /// </summary>
    private sealed class QueryBlockCollector : TSqlFragmentVisitor
    {
        private readonly string _sql;
        private readonly List<QueryBlock> _blocks = [];
        private readonly HashSet<(int StartOffset, int Length)> _seen = [];

        /// <summary>
        /// 収集元SQLを保持するvisitorを初期化
        /// </summary>
        private QueryBlockCollector(string sql)
        {
            _sql = sql;
        }

        /// <summary>
        /// ASTからサブクエリブロックを収集
        /// </summary>
        public static List<QueryBlock> Collect(string sql, TSqlFragment fragment)
        {
            var visitor = new QueryBlockCollector(sql);
            fragment.Accept(visitor);
            return visitor._blocks;
        }

        /// <summary>
        /// CTEのクエリ式を追加
        /// </summary>
        public override void ExplicitVisit(CommonTableExpression node)
        {
            AddQueryExpression(node.QueryExpression, "CTE");
            base.ExplicitVisit(node);
        }

        /// <summary>
        /// 派生テーブルのクエリ式を追加
        /// </summary>
        public override void ExplicitVisit(QueryDerivedTable node)
        {
            AddQueryExpression(node.QueryExpression, "SUBQUERY");
            base.ExplicitVisit(node);
        }

        /// <summary>
        /// スカラーサブクエリを追加
        /// </summary>
        public override void ExplicitVisit(ScalarSubquery node)
        {
            AddQueryExpression(node.QueryExpression, "SUBQUERY");
            base.ExplicitVisit(node);
        }

        /// <summary>
        /// EXISTSのサブクエリを追加
        /// </summary>
        public override void ExplicitVisit(ExistsPredicate node)
        {
            AddQueryExpression(node.Subquery.QueryExpression, "SUBQUERY");
            base.ExplicitVisit(node);
        }

        /// <summary>
        /// INのサブクエリを追加
        /// </summary>
        public override void ExplicitVisit(InPredicate node)
        {
            if (node.Subquery is not null)
            {
                AddQueryExpression(node.Subquery.QueryExpression, "SUBQUERY");
            }

            base.ExplicitVisit(node);
        }

        /// <summary>
        /// 子を先に収集してからクエリ式を重複なく追加
        /// </summary>
        private void AddQueryExpression(QueryExpression queryExpression, string kind)
        {
            queryExpression.Accept(this);

            var key = (queryExpression.StartOffset, queryExpression.FragmentLength);
            if (!_seen.Add(key))
            {
                return;
            }

            var text = FragmentText(_sql, queryExpression);
            if (text.Length == 0)
            {
                return;
            }

            _blocks.Add(new QueryBlock(kind, text, queryExpression.StartOffset, queryExpression.FragmentLength));
        }
    }
}
