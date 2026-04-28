namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetCellValue
    {
        public int Row { get; set; }

        public int Column { get; set; }

        public string Text { get; set; } = string.Empty;
    }
}
