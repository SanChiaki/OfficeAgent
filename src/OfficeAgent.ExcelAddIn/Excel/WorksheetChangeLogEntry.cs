using System;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetChangeLogEntry
    {
        public string Key { get; set; } = string.Empty;

        public string HeaderText { get; set; } = string.Empty;

        public string ChangeMode { get; set; } = string.Empty;

        public string NewValue { get; set; } = string.Empty;

        public string OldValue { get; set; } = string.Empty;

        public DateTime ChangedAt { get; set; }
    }
}
