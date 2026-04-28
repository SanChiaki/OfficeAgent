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

        public IDisposable BeginBulkOperation()
        {
            return new BulkOperationScope(application);
        }

        public void EnsureWorksheetExists(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            var activeSheet = application.ActiveSheet as ExcelInterop.Worksheet;
            try
            {
                var workbook = GetWorkbook();
                if (FindWorksheet(workbook, sheetName) != null)
                {
                    return;
                }

                var lastSheet = workbook.Worksheets[workbook.Worksheets.Count] as ExcelInterop.Worksheet;
                var worksheet = workbook.Worksheets.Add(After: lastSheet) as ExcelInterop.Worksheet;
                if (worksheet == null)
                {
                    throw new InvalidOperationException($"Worksheet '{sheetName}' could not be created.");
                }

                worksheet.Name = sheetName;
            }
            finally
            {
                try
                {
                    activeSheet?.Activate();
                }
                catch
                {
                    // Ignore focus restoration failures and keep worksheet creation successful.
                }
            }
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

            var rowCount = endRow - startRow + 1;
            var columnCount = endColumn - startColumn + 1;

            var worksheet = GetWorksheet(sheetName);
            var range = worksheet.Range[
                worksheet.Cells[startRow, startColumn],
                worksheet.Cells[endRow, endColumn]] as ExcelInterop.Range;
            if (range == null)
            {
                return new object[0, 0];
            }

            return NormalizeToObjectMatrix(range.Value2, rowCount, columnCount);
        }

        public string[,] ReadRangeNumberFormats(string sheetName, int startRow, int endRow, int startColumn, int endColumn)
        {
            if (endRow < startRow || endColumn < startColumn)
            {
                return new string[0, 0];
            }

            var rowCount = endRow - startRow + 1;
            var columnCount = endColumn - startColumn + 1;

            var worksheet = GetWorksheet(sheetName);
            var range = worksheet.Range[
                worksheet.Cells[startRow, startColumn],
                worksheet.Cells[endRow, endColumn]] as ExcelInterop.Range;
            if (range == null)
            {
                return new string[0, 0];
            }

            object rangeNumberFormat = range.NumberFormat;
            Func<int, int, object> readCellNumberFormat = (rowOffset, columnOffset) =>
            {
                var cell = worksheet.Cells[startRow + rowOffset, startColumn + columnOffset] as ExcelInterop.Range;
                return cell?.NumberFormat;
            };

            return NormalizeNumberFormatsWithFallback(
                rangeNumberFormat,
                rowCount,
                columnCount,
                readCellNumberFormat);
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

        private static object[,] NormalizeToObjectMatrix(object value, int requestedRowCount, int requestedColumnCount)
        {
            if (requestedRowCount <= 0 || requestedColumnCount <= 0)
            {
                return new object[0, 0];
            }

            var normalized = new object[requestedRowCount, requestedColumnCount];
            if (value is object[,] matrix)
            {
                var rowCount = Math.Min(matrix.GetLength(0), requestedRowCount);
                var columnCount = Math.Min(matrix.GetLength(1), requestedColumnCount);
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

            for (var row = 0; row < requestedRowCount; row++)
            {
                for (var column = 0; column < requestedColumnCount; column++)
                {
                    normalized[row, column] = value;
                }
            }

            return normalized;
        }

        private static string[,] NormalizeToStringMatrix(object value, int requestedRowCount, int requestedColumnCount)
        {
            if (requestedRowCount <= 0 || requestedColumnCount <= 0)
            {
                return new string[0, 0];
            }

            var normalized = new string[requestedRowCount, requestedColumnCount];
            if (value is object[,] matrix)
            {
                var rowCount = Math.Min(matrix.GetLength(0), requestedRowCount);
                var columnCount = Math.Min(matrix.GetLength(1), requestedColumnCount);
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

            var scalar = Convert.ToString(value) ?? string.Empty;
            for (var row = 0; row < requestedRowCount; row++)
            {
                for (var column = 0; column < requestedColumnCount; column++)
                {
                    normalized[row, column] = scalar;
                }
            }

            return normalized;
        }

        private static string[,] NormalizeNumberFormatsWithFallback(
            object rangeNumberFormat,
            int requestedRowCount,
            int requestedColumnCount,
            Func<int, int, object> readCellNumberFormat)
        {
            if (requestedRowCount <= 0 || requestedColumnCount <= 0)
            {
                return new string[0, 0];
            }

            if (rangeNumberFormat == null || rangeNumberFormat == DBNull.Value)
            {
                return ReadPerCellNumberFormats(requestedRowCount, requestedColumnCount, readCellNumberFormat);
            }

            if (rangeNumberFormat is object[,] matrix)
            {
                if (matrix.GetLength(0) < requestedRowCount || matrix.GetLength(1) < requestedColumnCount)
                {
                    return ReadPerCellNumberFormats(requestedRowCount, requestedColumnCount, readCellNumberFormat);
                }

                return NormalizeToStringMatrix(matrix, requestedRowCount, requestedColumnCount);
            }

            if (rangeNumberFormat is Array)
            {
                return ReadPerCellNumberFormats(requestedRowCount, requestedColumnCount, readCellNumberFormat);
            }

            return NormalizeToStringMatrix(rangeNumberFormat, requestedRowCount, requestedColumnCount);
        }

        private static string[,] ReadPerCellNumberFormats(
            int requestedRowCount,
            int requestedColumnCount,
            Func<int, int, object> readCellNumberFormat)
        {
            var formats = new string[requestedRowCount, requestedColumnCount];
            if (readCellNumberFormat == null)
            {
                return formats;
            }

            for (var rowOffset = 0; rowOffset < requestedRowCount; rowOffset++)
            {
                for (var columnOffset = 0; columnOffset < requestedColumnCount; columnOffset++)
                {
                    formats[rowOffset, columnOffset] = Convert.ToString(readCellNumberFormat(rowOffset, columnOffset)) ?? string.Empty;
                }
            }

            return formats;
        }

        private ExcelInterop.Worksheet GetWorksheet(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentException("Sheet name is required.", nameof(sheetName));
            }

            var worksheet = FindWorksheet(GetWorkbook(), sheetName);
            if (worksheet != null)
            {
                return worksheet;
            }

            throw new InvalidOperationException($"Worksheet '{sheetName}' was not found.");
        }

        private ExcelInterop.Workbook GetWorkbook()
        {
            var workbook = application.ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException("Excel workbook is not available.");
            }

            return workbook;
        }

        private static ExcelInterop.Worksheet FindWorksheet(ExcelInterop.Workbook workbook, string sheetName)
        {
            if (workbook == null)
            {
                return null;
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

            return null;
        }

        private sealed class BulkOperationScope : IDisposable
        {
            private readonly ExcelInterop.Application application;
            private readonly bool previousScreenUpdating;
            private readonly bool previousEnableEvents;
            private readonly ExcelInterop.XlCalculation previousCalculation;
            private bool disposed;

            public BulkOperationScope(ExcelInterop.Application application)
            {
                this.application = application ?? throw new ArgumentNullException(nameof(application));
                previousScreenUpdating = application.ScreenUpdating;
                previousEnableEvents = application.EnableEvents;
                previousCalculation = application.Calculation;

                application.ScreenUpdating = false;
                application.EnableEvents = false;
                application.Calculation = ExcelInterop.XlCalculation.xlCalculationManual;
            }

            public void Dispose()
            {
                if (disposed)
                {
                    return;
                }

                disposed = true;
                application.Calculation = previousCalculation;
                application.EnableEvents = previousEnableEvents;
                application.ScreenUpdating = previousScreenUpdating;
            }
        }
    }
}
