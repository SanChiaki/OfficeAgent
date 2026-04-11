namespace OfficeAgent.Core.Models
{
    public sealed class WorksheetSchema
    {
        public string SystemKey { get; set; } = string.Empty;
        public string ProjectId { get; set; } = string.Empty;
        public WorksheetColumnBinding[] Columns { get; set; } = System.Array.Empty<WorksheetColumnBinding>();
    }
}
