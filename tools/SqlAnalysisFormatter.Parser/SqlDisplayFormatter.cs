using Microsoft.SqlServer.TransactSql.ScriptDom;
using System.Text.RegularExpressions;

namespace SqlAnalysisFormatter.Parser;

/// <summary>
/// ASTで識別したSQL断片を帳票向けの表記へ整形
/// </summary>
internal static class SqlDisplayFormatter
{
    /// <summary>
    /// 文字列リテラルを維持したまま関数名とSQL演算子を大文字へ統一
    /// </summary>
    public static string Format(
        string sql,
        TSqlFragment fragment,
        bool uppercaseDateParts = false,
        bool compactUnarySigns = false)
    {
        if (fragment.StartOffset < 0 || fragment.FragmentLength <= 0 ||
            fragment.StartOffset + fragment.FragmentLength > sql.Length)
        {
            return string.Empty;
        }

        var text = sql.Substring(fragment.StartOffset, fragment.FragmentLength);
        var replacements = CollectReplacements(fragment);
        var characters = text.ToCharArray();
        foreach (var replacement in replacements)
        {
            var relativeOffset = replacement.Key - fragment.StartOffset;
            if (relativeOffset < 0 || relativeOffset + replacement.Value.Length > characters.Length)
            {
                continue;
            }

            replacement.Value.CopyTo(0, characters, relativeOffset, replacement.Value.Length);
        }

        text = new string(characters);
        if (uppercaseDateParts)
        {
            text = Regex.Replace(
                text,
                @"\b(DATEADD|DATEDIFF)\(\s*(year|quarter|month|dayofyear|day|week|weekday|hour|minute|second|millisecond|microsecond|nanosecond)\b",
                match => match.Groups[1].Value + "(" + match.Groups[2].Value.ToUpperInvariant(),
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        if (compactUnarySigns)
        {
            text = Regex.Replace(
                text,
                @"(?<![\w])([+-])\s+(?=\d)",
                "$1",
                RegexOptions.CultureInvariant);
        }

        return text.Trim();
    }

    /// <summary>
    /// トークンと関数ASTから同じ長さの置換文字列を収集
    /// </summary>
    private static IReadOnlyDictionary<int, string> CollectReplacements(TSqlFragment fragment)
    {
        var replacements = new Dictionary<int, string>();
        if (fragment.ScriptTokenStream is not null &&
            fragment.FirstTokenIndex >= 0 &&
            fragment.LastTokenIndex < fragment.ScriptTokenStream.Count)
        {
            for (var index = fragment.FirstTokenIndex; index <= fragment.LastTokenIndex; index++)
            {
                var token = fragment.ScriptTokenStream[index];
                if (ShouldUppercase(token.TokenType))
                {
                    replacements[token.Offset] = token.Text.ToUpperInvariant();
                }
            }
        }

        fragment.Accept(new FunctionNameCollector(replacements));
        return replacements;
    }

    /// <summary>
    /// 条件式で大文字へ統一するSQLトークンを判定
    /// </summary>
    private static bool ShouldUppercase(TSqlTokenType tokenType)
    {
        return tokenType is
            TSqlTokenType.And or
            TSqlTokenType.Between or
            TSqlTokenType.Case or
            TSqlTokenType.Else or
            TSqlTokenType.End or
            TSqlTokenType.Exists or
            TSqlTokenType.In or
            TSqlTokenType.Is or
            TSqlTokenType.Like or
            TSqlTokenType.Not or
            TSqlTokenType.Null or
            TSqlTokenType.Or or
            TSqlTokenType.Then or
            TSqlTokenType.When;
    }

    /// <summary>
    /// 関数名の開始位置と大文字表記を収集
    /// </summary>
    private sealed class FunctionNameCollector(IDictionary<int, string> replacements) : TSqlFragmentVisitor
    {
        /// <summary>
        /// 通常関数の識別子を追加
        /// </summary>
        public override void ExplicitVisit(FunctionCall node)
        {
            replacements[node.FunctionName.StartOffset] = node.FunctionName.Value.ToUpperInvariant();
            base.ExplicitVisit(node);
        }

        /// <summary>
        /// COALESCE専用ASTの開始位置を追加
        /// </summary>
        public override void ExplicitVisit(CoalesceExpression node)
        {
            replacements[node.StartOffset] = "COALESCE";
            base.ExplicitVisit(node);
        }

        /// <summary>
        /// IIF専用ASTの開始位置を追加
        /// </summary>
        public override void ExplicitVisit(IIfCall node)
        {
            replacements[node.StartOffset] = "IIF";
            base.ExplicitVisit(node);
        }
    }
}
