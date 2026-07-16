using System.IO.Compression;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Xml.Linq;
using SqlAnalysisFormatter.Parser;

namespace SqlAnalysisFormatter.Parser.Tests;

/// <summary>
/// 登録済みExcel期待値67件との回帰テスト
/// </summary>
[TestClass]
public sealed class OutputGoldenRegressionTests
{
    /// <summary>
    /// 全67ケースが登録済み期待値と一致することを確認
    /// </summary>
    [TestMethod]
    public void Build_MatchesAllRegisteredExpectations()
    {
        var fixture = LoadFixture();
        using var workbook = new OutputExpectationWorkbook(
            Path.Combine(AppContext.BaseDirectory, "SqlAnalysisFormatter.OutputExpectations.xlsx"));
        var failures = new List<string>();

        foreach (var testCase in fixture.Cases)
        {
            var mappings = testCase.Tables
                .Select(table => new MappingDefinition(table.Key, table.Value, string.Empty, string.Empty))
                .Concat((testCase.OutputFields ?? new Dictionary<string, string>())
                    .Select(field => new MappingDefinition("-", string.Empty, field.Key, field.Value)))
                .ToArray();
            var plan = OutputSheetPlanBuilder.Build(string.Join('\n', testCase.SqlLines), mappings);
            var expected = workbook.ReadCells(testCase.Id);
            var actual = plan.Cells.ToDictionary(cell => (cell.Row, cell.Column), cell => cell.Value);
            failures.AddRange(CompareCase(testCase.Id, plan.RowCount, expected, actual));
        }

        Assert.IsEmpty(failures, string.Join(Environment.NewLine, failures));
    }

    /// <summary>
    /// JSONから和名変換済みSQLケースを読み込む
    /// </summary>
    private static OutputFixture LoadFixture()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "OutputReportCases.json");
        var fixture = JsonSerializer.Deserialize<OutputFixture>(
            File.ReadAllText(path),
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        return fixture ?? throw new InvalidDataException("OutputReportCases.jsonを読み込めません。");
    }

    /// <summary>
    /// 1ケースの行数と非空セルを比較
    /// </summary>
    private static IEnumerable<string> CompareCase(
        string caseId,
        int actualRowCount,
        IReadOnlyDictionary<(int Row, int Column), string> expected,
        IReadOnlyDictionary<(int Row, int Column), string> actual)
    {
        var expectedRowCount = expected.Keys.Max(key => key.Row);
        if (actualRowCount != expectedRowCount)
        {
            yield return $"{caseId}: 行数 expected={expectedRowCount}, actual={actualRowCount}";
        }

        foreach (var key in expected.Keys.Union(actual.Keys).OrderBy(key => key.Row).ThenBy(key => key.Column))
        {
            expected.TryGetValue(key, out var expectedValue);
            actual.TryGetValue(key, out var actualValue);
            if (!string.Equals(expectedValue, actualValue, StringComparison.Ordinal))
            {
                yield return $"{caseId} {CellAddress(key.Row, key.Column)}: expected=[{expectedValue}], actual=[{actualValue}]";
            }
        }
    }

    /// <summary>
    /// 行列番号をExcelセル番地へ変換
    /// </summary>
    private static string CellAddress(int row, int column)
    {
        var columnName = string.Empty;
        while (column > 0)
        {
            column--;
            columnName = (char)('A' + column % 26) + columnName;
            column /= 26;
        }

        return columnName + row;
    }

    private sealed record OutputFixture(IReadOnlyList<OutputFixtureCase> Cases);

    private sealed record OutputFixtureCase(
        string Id,
        [property: JsonPropertyName("sql_lines")]
        IReadOnlyList<string> SqlLines,
        IReadOnlyDictionary<string, string> Tables,
        [property: JsonPropertyName("output_fields")]
        IReadOnlyDictionary<string, string>? OutputFields);

    /// <summary>
    /// Open XMLパッケージから期待値セルだけを読み取る
    /// </summary>
    private sealed class OutputExpectationWorkbook : IDisposable
    {
        private static readonly XNamespace SpreadsheetNamespace =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace OfficeRelationshipNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace PackageRelationshipNamespace =
            "http://schemas.openxmlformats.org/package/2006/relationships";

        private readonly FileStream _stream;
        private readonly ZipArchive _archive;
        private readonly IReadOnlyDictionary<string, string> _sheetPaths;
        private readonly IReadOnlyList<string> _sharedStrings;

        /// <summary>
        /// 期待値ブックを読み取り専用で開く
        /// </summary>
        public OutputExpectationWorkbook(string path)
        {
            _stream = File.OpenRead(path);
            _archive = new ZipArchive(_stream, ZipArchiveMode.Read, leaveOpen: false);
            _sheetPaths = LoadSheetPaths();
            _sharedStrings = LoadSharedStrings();
        }

        /// <summary>
        /// 指定シートの非空セルを取得
        /// </summary>
        public IReadOnlyDictionary<(int Row, int Column), string> ReadCells(string sheetName)
        {
            if (!_sheetPaths.TryGetValue(sheetName, out var sheetPath))
            {
                throw new InvalidDataException($"期待値シートがありません: {sheetName}");
            }

            var document = LoadXml(sheetPath);
            var result = new Dictionary<(int Row, int Column), string>();
            foreach (var cell in document.Descendants(SpreadsheetNamespace + "c"))
            {
                var address = (string?)cell.Attribute("r") ?? string.Empty;
                var value = ReadCellValue(cell);
                if (address.Length == 0 || value.Length == 0)
                {
                    continue;
                }

                result[ParseAddress(address)] = value;
            }

            return result;
        }

        /// <summary>
        /// ZIPリソースを解放
        /// </summary>
        public void Dispose()
        {
            _archive.Dispose();
            _stream.Dispose();
        }

        /// <summary>
        /// シート名からワークシートXMLへの対応を作る
        /// </summary>
        private IReadOnlyDictionary<string, string> LoadSheetPaths()
        {
            var workbook = LoadXml("xl/workbook.xml");
            var relationships = LoadXml("xl/_rels/workbook.xml.rels")
                .Descendants(PackageRelationshipNamespace + "Relationship")
                .ToDictionary(
                    element => (string)element.Attribute("Id")!,
                    element => NormalizePartPath((string)element.Attribute("Target")!));

            return workbook.Descendants(SpreadsheetNamespace + "sheet")
                .ToDictionary(
                    sheet => (string)sheet.Attribute("name")!,
                    sheet => relationships[(string)sheet.Attribute(OfficeRelationshipNamespace + "id")!]);
        }

        /// <summary>
        /// 共有文字列テーブルを読み込む
        /// </summary>
        private IReadOnlyList<string> LoadSharedStrings()
        {
            if (_archive.GetEntry("xl/sharedStrings.xml") is null)
            {
                return [];
            }

            return LoadXml("xl/sharedStrings.xml")
                .Descendants(SpreadsheetNamespace + "si")
                .Select(ReadRichText)
                .ToArray();
        }

        /// <summary>
        /// ふりがな要素を除外してセル本文だけを連結
        /// </summary>
        private static string ReadRichText(XElement container)
        {
            var directText = container.Elements(SpreadsheetNamespace + "t");
            var runText = container.Elements(SpreadsheetNamespace + "r")
                .SelectMany(run => run.Elements(SpreadsheetNamespace + "t"));
            return string.Concat(directText.Concat(runText).Select(text => text.Value));
        }

        /// <summary>
        /// セル型に応じて表示値を取得
        /// </summary>
        private string ReadCellValue(XElement cell)
        {
            var type = (string?)cell.Attribute("t");
            if (type == "inlineStr")
            {
                var inlineString = cell.Element(SpreadsheetNamespace + "is");
                return inlineString is null ? string.Empty : ReadRichText(inlineString);
            }

            var rawValue = cell.Element(SpreadsheetNamespace + "v")?.Value ?? string.Empty;
            return type == "s" && int.TryParse(rawValue, out var sharedIndex)
                ? _sharedStrings[sharedIndex]
                : rawValue;
        }

        /// <summary>
        /// セル番地を行列番号へ分解
        /// </summary>
        private static (int Row, int Column) ParseAddress(string address)
        {
            var index = 0;
            var column = 0;
            while (index < address.Length && char.IsLetter(address[index]))
            {
                column = column * 26 + char.ToUpperInvariant(address[index]) - 'A' + 1;
                index++;
            }

            return (int.Parse(address[index..]), column);
        }

        /// <summary>
        /// 関係ファイルの相対パスをZIP内パスへ変換
        /// </summary>
        private static string NormalizePartPath(string target)
        {
            target = target.Replace('\\', '/');
            return target.StartsWith('/') ? target.TrimStart('/') : "xl/" + target;
        }

        /// <summary>
        /// ZIP内XMLを読み込む
        /// </summary>
        private XDocument LoadXml(string path)
        {
            var entry = _archive.GetEntry(path)
                ?? throw new InvalidDataException($"Open XMLパーツがありません: {path}");
            using var stream = entry.Open();
            return XDocument.Load(stream);
        }
    }
}
