namespace OfficeAgent.Core.Models
{
    public sealed class SelectedVisibleCell
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string Value { get; set; } = string.Empty;
    }
}
