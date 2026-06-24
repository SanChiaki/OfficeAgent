using System;

namespace OfficeAgent.Core.Models
{
    public sealed class BusinessExportWorkbook
    {
        public string FileName { get; set; } = string.Empty;

        public string ContentType { get; set; } = string.Empty;

        public byte[] Content { get; set; } = Array.Empty<byte>();
    }
}
