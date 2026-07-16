using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace SqlAnalysisFormatter.Parser;

/// <summary>
/// VBAから呼び出すSQL解析CLI
/// </summary>
internal static class Program
{
    private const string VersionText = "SqlAnalysisFormatter.Parser 0.4.0";

    /// <summary>
    /// SQLを読み取り、指定形式の解析結果を書き出す
    /// </summary>
    private static async Task<int> Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

        if (args.Length == 1 && args[0] == "--version")
        {
            Console.WriteLine(VersionText);
            return 0;
        }

        string sql;
        CliOptions options;
        try
        {
            options = ParseOptions(args);
            sql = await ReadSqlAsync(options);
        }
        catch (ArgumentException ex)
        {
            Console.Error.WriteLine(ex.Message);
            return 2;
        }

        try
        {
            if (options.Format == "vba-plan")
            {
                var mappings = await ReadMappingsAsync(options.MappingPath);
                var plan = OutputSheetPlanBuilder.Build(sql, mappings);
                await WriteTextAsync(VbaOutputProtocol.SerializePlan(plan), options, Encoding.Unicode);
                return 0;
            }

            var result = TsqlAstParser.Parse(sql);
            await WriteResultAsync(result, options);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or InvalidDataException)
        {
            Console.Error.WriteLine(ex.Message);
            return 2;
        }

        return 0;
    }

    /// <summary>
    /// コマンドライン引数を検証してオプションへ変換
    /// </summary>
    private static CliOptions ParseOptions(string[] args)
    {
        var options = new CliOptions();
        for (var index = 0; index < args.Length; index++)
        {
            switch (args[index])
            {
                case "--input":
                    options.InputPath = ReadOptionValue(args, ref index, "--input");
                    break;
                case "--output":
                    options.OutputPath = ReadOptionValue(args, ref index, "--output");
                    break;
                case "--mappings":
                    options.MappingPath = ReadOptionValue(args, ref index, "--mappings");
                    break;
                case "--format":
                    options.Format = ReadOptionValue(args, ref index, "--format");
                    break;
                default:
                    throw new ArgumentException("Usage: SqlAnalysisFormatter.Parser.exe [--input <path>] [--output <path>] [--mappings <path>] [--format json|vba-blocks|vba-plan] [--version]");
            }
        }

        if (options.Format is not "json" and not "vba-blocks" and not "vba-plan")
        {
            throw new ArgumentException("Unknown format: " + options.Format);
        }

        return options;
    }

    /// <summary>
    /// 値を伴うオプションの次要素を取得
    /// </summary>
    private static string ReadOptionValue(string[] args, ref int index, string optionName)
    {
        if (index + 1 >= args.Length)
        {
            throw new ArgumentException("Missing value for " + optionName);
        }

        index++;
        return args[index];
    }

    /// <summary>
    /// ファイルまたは標準入力からSQLを読み込む
    /// </summary>
    private static async Task<string> ReadSqlAsync(CliOptions options)
    {
        if (options.InputPath.Length == 0)
        {
            using var reader = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8);
            return await reader.ReadToEndAsync();
        }

        return await File.ReadAllTextAsync(options.InputPath, Encoding.UTF8);
    }

    /// <summary>
    /// 従来のJSONまたはブロック形式を書き出す
    /// </summary>
    private static async Task WriteResultAsync(ParseResult result, CliOptions options)
    {
        var outputText = options.Format == "vba-blocks"
            ? ParserOutputFormatter.ToVbaBlocks(result)
            : JsonSerializer.Serialize(result, JsonOptions());

        var encoding = options.Format == "vba-blocks"
            ? Encoding.Unicode
            : new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        await WriteTextAsync(outputText, options, encoding);
    }

    /// <summary>
    /// VBA連携用ファイルから変換定義を読み込む
    /// </summary>
    private static async Task<IReadOnlyList<MappingDefinition>> ReadMappingsAsync(string mappingPath)
    {
        if (mappingPath.Length == 0)
        {
            return [];
        }

        var text = await File.ReadAllTextAsync(mappingPath, Encoding.UTF8);
        return VbaOutputProtocol.ParseMappings(text);
    }

    /// <summary>
    /// ファイルまたは標準出力へ指定エンコーディングで書き出す
    /// </summary>
    private static async Task WriteTextAsync(string outputText, CliOptions options, Encoding encoding)
    {
        if (options.OutputPath.Length == 0)
        {
            Console.WriteLine(outputText);
            return;
        }

        await File.WriteAllTextAsync(options.OutputPath, outputText, encoding);
    }

    /// <summary>
    /// 日本語をエスケープしないJSON設定を作成
    /// </summary>
    private static JsonSerializerOptions JsonOptions()
    {
        return new JsonSerializerOptions
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            WriteIndented = false
        };
    }

    private sealed class CliOptions
    {
        public string InputPath { get; set; } = string.Empty;

        public string OutputPath { get; set; } = string.Empty;

        public string MappingPath { get; set; } = string.Empty;

        public string Format { get; set; } = "json";
    }
}
