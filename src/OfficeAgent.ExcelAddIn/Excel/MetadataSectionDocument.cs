using System;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class MetadataSectionDocument
    {
        public MetadataSectionDocument(string title, string[] headers, string[][] rows)
        {
            Title = title ?? string.Empty;
            Headers = headers ?? Array.Empty<string>();
            Rows = rows ?? Array.Empty<string[]>();
        }

        public string Title { get; }

        public string[] Headers { get; }

        public string[][] Rows { get; }
    }
}
