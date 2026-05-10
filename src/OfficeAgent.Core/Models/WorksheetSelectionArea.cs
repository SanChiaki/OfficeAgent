namespace OfficeAgent.Core.Models
{
    public sealed class WorksheetSelectionArea
    {
        public int StartRow { get; set; }

        public int EndRow { get; set; }

        public int StartColumn { get; set; }

        public int EndColumn { get; set; }
    }
}
