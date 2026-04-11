namespace OfficeAgent.Core.Models
{
    public enum WorksheetColumnKind
    {
        Single,
        ActivityProperty,
    }

    public sealed class WorksheetColumnBinding
    {
        public int ColumnIndex { get; set; }
        public string ApiFieldKey { get; set; } = string.Empty;
        public WorksheetColumnKind ColumnKind { get; set; }
        public string ParentHeaderText { get; set; } = string.Empty;
        public string ChildHeaderText { get; set; } = string.Empty;
        public string ActivityId { get; set; } = string.Empty;
        public string ActivityName { get; set; } = string.Empty;
        public string PropertyKey { get; set; } = string.Empty;
        public bool IsIdColumn { get; set; }
    }
}
