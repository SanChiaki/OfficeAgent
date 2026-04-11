namespace OfficeAgent.Core.Models
{
    public sealed class HeaderCellPlan
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public int RowSpan { get; set; } = 1;
        public int ColumnSpan { get; set; } = 1;
        public string Text { get; set; } = string.Empty;
    }
}
