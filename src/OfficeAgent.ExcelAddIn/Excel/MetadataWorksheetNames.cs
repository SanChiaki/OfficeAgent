using System;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal static class MetadataWorksheetNames
    {
        public const string Current = "xISDP_Setting";
        public const string Legacy = "ISDP_Setting";

        public static bool IsMetadataWorksheet(string sheetName)
        {
            return string.Equals(sheetName, Current, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(sheetName, Legacy, StringComparison.OrdinalIgnoreCase);
        }
    }
}
