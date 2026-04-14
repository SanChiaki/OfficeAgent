namespace OfficeAgent.Core.Models
{
    public sealed class WorksheetRuntimeColumn
    {
        public int ColumnIndex { get; set; }

        public string ApiFieldKey { get; set; } = string.Empty;

        public string HeaderType { get; set; } = string.Empty;

        public string DisplayText { get; set; } = string.Empty;

        public string ParentDisplayText { get; set; } = string.Empty;

        public string ChildDisplayText { get; set; } = string.Empty;

        public bool IsIdColumn { get; set; }
    }
}
