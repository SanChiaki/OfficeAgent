namespace OfficeAgent.Core.Models
{
    public sealed class CellChange
    {
        public string SheetName { get; set; } = string.Empty;
        public string RowId { get; set; } = string.Empty;
        public string ApiFieldKey { get; set; } = string.Empty;
        public string OldValue { get; set; } = string.Empty;
        public string NewValue { get; set; } = string.Empty;
    }
}
