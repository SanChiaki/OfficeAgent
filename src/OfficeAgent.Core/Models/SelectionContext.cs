using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace OfficeAgent.Core.Models
{
    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class SelectionContext
    {
        public bool HasSelection { get; set; }

        public string WorkbookName { get; set; } = string.Empty;

        public string SheetName { get; set; } = string.Empty;

        public string Address { get; set; } = string.Empty;

        public int RowCount { get; set; }

        public int ColumnCount { get; set; }

        public bool IsContiguous { get; set; } = true;

        public string[] HeaderPreview { get; set; } = System.Array.Empty<string>();

        public string[][] SampleRows { get; set; } = System.Array.Empty<string[]>();

        public string WarningMessage { get; set; }

        public static SelectionContext Empty(string warningMessage)
        {
            return new SelectionContext
            {
                HasSelection = false,
                WarningMessage = warningMessage,
            };
        }
    }
}
