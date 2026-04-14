namespace OfficeAgent.Core.Models
{
    public sealed class SheetBinding
    {
        public string SheetName { get; set; } = string.Empty;
        public string SystemKey { get; set; } = string.Empty;
        public string ProjectId { get; set; } = string.Empty;
        public string ProjectName { get; set; } = string.Empty;
        public int HeaderStartRow { get; set; } = 1;
        public int HeaderRowCount { get; set; } = 2;
        public int DataStartRow { get; set; } = 3;
    }
}
