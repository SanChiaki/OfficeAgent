using System;

namespace OfficeAgent.Core.Models
{
    public sealed class SyncOperationPreview
    {
        public string OperationName { get; set; } = string.Empty;
        public string Summary { get; set; } = string.Empty;
        public string[] Details { get; set; } = Array.Empty<string>();
        public CellChange[] Changes { get; set; } = Array.Empty<CellChange>();
        public SkippedCellChange[] SkippedChanges { get; set; } = Array.Empty<SkippedCellChange>();
    }
}
