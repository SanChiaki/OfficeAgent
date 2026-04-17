using System;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class ExcelWorksheetGridAdapter : IWorksheetGridAdapter
    {
        private readonly ExcelInterop.Application application;
        private static readonly IDisposable NoOpBulkOperationScope = new NoOpScope();

        public ExcelWorksheetGridAdapter(ExcelInterop.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public IDisposable BeginBulkOperation()
        {
            return NoOpBulkOperationScope;
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

        public void WriteRangeValues(string sheetName, int startRow, int startColumn, object[,] values)
        {
            if (values == null)
            {
                throw new ArgumentNullException(nameof(values));
            }

            var rowCount = values.GetLength(0);
            var columnCount = values.GetLength(1);
            if (rowCount <= 0 || columnCount <= 0)
            {
                return;
            }

            var worksheet = GetWorksheet(sheetName);
            var startCell = worksheet.Cells[startRow, startColumn] as ExcelInterop.Range;
            var range = startCell?.Resize[rowCount, columnCount] as ExcelInterop.Range;
            if (range == null)
            {
                return;
            }

            range.Value2 = values;
        }

        public object[,] ReadRangeValues(string sheetName, int startRow, int endRow, int startColumn, int endColumn)
        {
            if (endRow < startRow || endColumn < startColumn)
            {
                return new object[0, 0];
            }

            var worksheet = GetWorksheet(sheetName);
            var range = worksheet.Range[
                worksheet.Cells[startRow, startColumn],
                worksheet.Cells[endRow, endColumn]] as ExcelInterop.Range;
            if (range == null)
            {
                return new object[0, 0];
            }

            return NormalizeToObjectMatrix(range.Value2);
        }

        public string[,] ReadRangeNumberFormats(string sheetName, int startRow, int endRow, int startColumn, int endColumn)
        {
            if (endRow < startRow || endColumn < startColumn)
            {
                return new string[0, 0];
            }

            var worksheet = GetWorksheet(sheetName);
            var range = worksheet.Range[
                worksheet.Cells[startRow, startColumn],
                worksheet.Cells[endRow, endColumn]] as ExcelInterop.Range;
            if (range == null)
            {
                return new string[0, 0];
            }

            return NormalizeToStringMatrix(range.NumberFormat);
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

        private static object[,] NormalizeToObjectMatrix(object value)
        {
            if (value is object[,] matrix)
            {
                var rowCount = matrix.GetLength(0);
                var columnCount = matrix.GetLength(1);
                var normalized = new object[rowCount, columnCount];
                var rowLowerBound = matrix.GetLowerBound(0);
                var columnLowerBound = matrix.GetLowerBound(1);

                for (var row = 0; row < rowCount; row++)
                {
                    for (var column = 0; column < columnCount; column++)
                    {
                        normalized[row, column] = matrix[row + rowLowerBound, column + columnLowerBound];
                    }
                }

                return normalized;
            }

            return new object[1, 1] { { value } };
        }

        private static string[,] NormalizeToStringMatrix(object value)
        {
            if (value is object[,] matrix)
            {
                var rowCount = matrix.GetLength(0);
                var columnCount = matrix.GetLength(1);
                var normalized = new string[rowCount, columnCount];
                var rowLowerBound = matrix.GetLowerBound(0);
                var columnLowerBound = matrix.GetLowerBound(1);

                for (var row = 0; row < rowCount; row++)
                {
                    for (var column = 0; column < columnCount; column++)
                    {
                        normalized[row, column] = Convert.ToString(matrix[row + rowLowerBound, column + columnLowerBound]) ?? string.Empty;
                    }
                }

                return normalized;
            }

            return new string[1, 1] { { Convert.ToString(value) ?? string.Empty } };
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

        private sealed class NoOpScope : IDisposable
        {
            public void Dispose()
            {
            }
        }
    }
}
