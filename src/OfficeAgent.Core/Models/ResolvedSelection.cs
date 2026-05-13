namespace OfficeAgent.Core.Models
{
    public sealed class ResolvedSelection
    {
        public string[] RowIds { get; set; } = System.Array.Empty<string>();
        public string[] ApiFieldKeys { get; set; } = System.Array.Empty<string>();
        public SelectedVisibleCell[] TargetCells { get; set; } = System.Array.Empty<SelectedVisibleCell>();
        public WorksheetSelectionRow[] TargetRows { get; set; } = System.Array.Empty<WorksheetSelectionRow>();
        public int[] TargetColumns { get; set; } = System.Array.Empty<int>();
        public WorksheetSelectionArea[] TargetAreas { get; set; } = System.Array.Empty<WorksheetSelectionArea>();
    }
}
