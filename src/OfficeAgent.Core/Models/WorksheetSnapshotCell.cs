namespace OfficeAgent.Core.Models
{
    public sealed class WorksheetSnapshotCell
    {
        public string SheetName { get; set; } = string.Empty;
        public string RowId { get; set; } = string.Empty;
        public string ApiFieldKey { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }
}
