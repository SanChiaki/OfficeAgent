using System;

namespace OfficeAgent.Core.Models
{
    public sealed class UploadChangeFilterResult
    {
        public CellChange[] IncludedChanges { get; set; } = Array.Empty<CellChange>();

        public SkippedCellChange[] SkippedChanges { get; set; } = Array.Empty<SkippedCellChange>();
    }
}
