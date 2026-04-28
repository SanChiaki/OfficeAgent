namespace OfficeAgent.Core.Models
{
    public sealed class SkippedCellChange
    {
        public CellChange Change { get; set; } = new CellChange();

        public string Reason { get; set; } = string.Empty;
    }
}
