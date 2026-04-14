using System;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class ExcelWorksheetGridAdapter : IWorksheetGridAdapter
    {
        private readonly ExcelInterop.Application application;

        public ExcelWorksheetGridAdapter(ExcelInterop.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public string GetCellText(string sheetName, int row, int column)
        {
            var worksheet = GetWorksheet(sheetName);
            var cell = worksheet.Cells[row, column] as ExcelInterop.Range;
            return Convert.ToString(cell?.Text) ?? string.Empty;
        }

        public void SetCellText(string sheetName, int row, int column, string value)
        {
            var worksheet = GetWorksheet(sheetName);
            var cell = worksheet.Cells[row, column] as ExcelInterop.Range;
            cell.Value2 = value ?? string.Empty;
        }

        public void ClearRange(string sheetName, int startRow, int endRow, int startColumn, int endColumn)
        {
            if (endRow < startRow || endColumn < startColumn)
            {
                return;
            }

            var worksheet = GetWorksheet(sheetName);
            var range = worksheet.Range[
                worksheet.Cells[startRow, startColumn],
                worksheet.Cells[endRow, endColumn]] as ExcelInterop.Range;
            ClearRange(range);
        }

        public void ClearWorksheet(string sheetName)
        {
            var worksheet = GetWorksheet(sheetName);
            var usedRange = worksheet.UsedRange;
            ClearRange(usedRange);
        }

        public void MergeCells(string sheetName, int row, int column, int rowSpan, int columnSpan)
        {
            if (rowSpan <= 1 && columnSpan <= 1)
            {
                return;
            }

            var worksheet = GetWorksheet(sheetName);
            var range = worksheet.Range[
                worksheet.Cells[row, column],
                worksheet.Cells[row + rowSpan - 1, column + columnSpan - 1]];
            range.Merge();
        }

        public int GetLastUsedRow(string sheetName)
        {
            var worksheet = GetWorksheet(sheetName);
            var usedRange = worksheet.UsedRange;
            if (usedRange == null || usedRange.Rows == null || usedRange.Rows.Count == 0)
            {
                return 0;
            }

            return usedRange.Row + usedRange.Rows.Count - 1;
        }

        public int GetLastUsedColumn(string sheetName)
        {
            var worksheet = GetWorksheet(sheetName);
            var usedRange = worksheet.UsedRange;
            if (usedRange == null || usedRange.Columns == null || usedRange.Columns.Count == 0)
            {
                return 0;
            }

            return usedRange.Column + usedRange.Columns.Count - 1;
        }

        private static void ClearRange(ExcelInterop.Range range)
        {
            if (range == null)
            {
                return;
            }

            try
            {
                range.UnMerge();
            }
            catch
            {
                // Ignore when the range has no merged cells.
            }

            range.Clear();
        }

        private ExcelInterop.Worksheet GetWorksheet(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            var workbook = application.ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException("Excel workbook is not available.");
            }

            for (var index = 1; index <= workbook.Worksheets.Count; index++)
            {
                var worksheet = workbook.Worksheets[index] as ExcelInterop.Worksheet;
                if (worksheet != null &&
                    string.Equals(worksheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return worksheet;
                }
            }

            throw new InvalidOperationException($"Worksheet '{sheetName}' was not found.");
        }
    }
}
