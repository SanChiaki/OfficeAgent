using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace OfficeAgent.Core.Models
{
    public static class ExcelCommandTypes
    {
        public const string ReadSelectionTable = "excel.readSelectionTable";
        public const string ReadRange = "excel.readRange";
        public const string WriteRange = "excel.writeRange";
        public const string AddWorksheet = "excel.addWorksheet";
        public const string RenameWorksheet = "excel.renameWorksheet";
        public const string DeleteWorksheet = "excel.deleteWorksheet";
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class ExcelCommand
    {
        public string CommandType { get; set; } = string.Empty;

        public string SheetName { get; set; } = string.Empty;

        public string TargetAddress { get; set; } = string.Empty;

        public string NewSheetName { get; set; } = string.Empty;

        public string[][] Values { get; set; } = System.Array.Empty<string[]>();

        public bool Confirmed { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class ExcelCommandPreview
    {
        public string Title { get; set; } = string.Empty;

        public string Summary { get; set; } = string.Empty;

        public string[] Details { get; set; } = System.Array.Empty<string>();
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class ExcelTableData
    {
        public string SheetName { get; set; } = string.Empty;

        public string Address { get; set; } = string.Empty;

        public string[] Headers { get; set; } = System.Array.Empty<string>();

        public string[][] Rows { get; set; } = System.Array.Empty<string[]>();
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class ExcelCommandResult
    {
        public string CommandType { get; set; } = string.Empty;

        public bool RequiresConfirmation { get; set; }

        public string Status { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;

        public ExcelCommandPreview Preview { get; set; }

        public ExcelTableData Table { get; set; }

        public SelectionContext SelectionContext { get; set; }
    }
}
