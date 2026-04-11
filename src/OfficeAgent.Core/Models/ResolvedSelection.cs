namespace OfficeAgent.Core.Models
{
    public sealed class ResolvedSelection
    {
        public string[] RowIds { get; set; } = System.Array.Empty<string>();
        public string[] ApiFieldKeys { get; set; } = System.Array.Empty<string>();
        public SelectedVisibleCell[] TargetCells { get; set; } = System.Array.Empty<SelectedVisibleCell>();
    }
}
