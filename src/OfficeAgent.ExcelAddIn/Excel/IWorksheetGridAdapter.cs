namespace OfficeAgent.ExcelAddIn.Excel
{
    internal interface IWorksheetGridAdapter
    {
        string GetCellText(string sheetName, int row, int column);

        void SetCellText(string sheetName, int row, int column, string value);

        void ClearRange(string sheetName, int startRow, int endRow, int startColumn, int endColumn);

        void ClearWorksheet(string sheetName);

        void MergeCells(string sheetName, int row, int column, int rowSpan, int columnSpan);

        int GetLastUsedRow(string sheetName);

        int GetLastUsedColumn(string sheetName);
    }
}
