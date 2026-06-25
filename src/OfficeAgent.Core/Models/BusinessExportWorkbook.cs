using System;

namespace OfficeAgent.Core.Models
{
    public sealed class BusinessExportWorkbook
    {
        public string FileName { get; set; } = "business-export.xlsx";

        public string ContentType { get; set; } = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        public byte[] Content { get; set; } = Array.Empty<byte>();
    }
}
