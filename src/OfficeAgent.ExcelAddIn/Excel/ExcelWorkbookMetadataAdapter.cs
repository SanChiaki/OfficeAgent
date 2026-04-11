using System;
using System.Collections.Generic;
using System.Linq;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class ExcelWorkbookMetadataAdapter : IWorksheetMetadataAdapter
    {
        private const string MetadataSheetName = "_OfficeAgentMetadata";
        private readonly ExcelInterop.Application application;

        public ExcelWorkbookMetadataAdapter(ExcelInterop.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public void EnsureWorksheet(string name, bool visible)
        {
            var worksheet = EnsureWorksheetExists(name);
            worksheet.Visible = visible
                ? ExcelInterop.XlSheetVisibility.xlSheetVisible
                : ExcelInterop.XlSheetVisibility.xlSheetHidden;
        }

        public void WriteTable(string tableName, string[] headers, string[][] rows)
        {
            if (string.IsNullOrWhiteSpace(tableName))
            {
                throw new ArgumentException("Table name is required.", nameof(tableName));
            }

            if (headers == null)
            {
                throw new ArgumentNullException(nameof(headers));
            }

            if (rows == null)
            {
                throw new ArgumentNullException(nameof(rows));
            }

            var worksheet = EnsureWorksheetExists(MetadataSheetName);
            var dataColumns = Math.Max(headers.Length, rows.Length > 0 ? rows.Max(row => row?.Length ?? 0) : 0);
            var columnsToClear = 1 + Math.Max(dataColumns, 0);
            ClearTableRows(worksheet, tableName, columnsToClear);

            var startRow = GetFirstEmptyRow(worksheet);
            for (var i = 0; i < rows.Length; i++)
            {
                var values = rows[i] ?? Array.Empty<string>();
                WriteRow(worksheet, startRow + i, tableName, values, dataColumns);
            }
        }

        public string[][] ReadTable(string tableName)
        {
            if (string.IsNullOrWhiteSpace(tableName))
            {
                throw new ArgumentException("Table name is required.", nameof(tableName));
            }

            var worksheet = EnsureWorksheetExists(MetadataSheetName);
            var usedRange = worksheet.UsedRange;
            if (usedRange == null || usedRange.Rows.Count == 0)
            {
                return Array.Empty<string[]>();
            }

            var dataColumns = Math.Max(0, usedRange.Columns.Count - 1);
            var results = new List<string[]>();
            var startRow = usedRange.Row;
            var endRow = startRow + usedRange.Rows.Count - 1;

            for (var row = startRow; row <= endRow; row++)
            {
                var tableCell = worksheet.Cells[row, 1] as ExcelInterop.Range;
                if (tableCell?.Value2 is not string storedTable)
                {
                    continue;
                }

                if (!string.Equals(storedTable, tableName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var columnCount = DetermineActualColumnCount(worksheet, row, dataColumns);
                var values = new string[columnCount];

                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    var cell = worksheet.Cells[row, columnIndex + 2] as ExcelInterop.Range;
                    values[columnIndex] = cell?.Value2 as string ?? string.Empty;
                }

                results.Add(values);
            }

            return results.ToArray();
        }

        private ExcelInterop.Worksheet EnsureWorksheetExists(string name)
        {
            var workbook = GetWorkbook();

            foreach (ExcelInterop.Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }

            var lastSheet = workbook.Worksheets[workbook.Worksheets.Count] as ExcelInterop.Worksheet;
            var worksheet = workbook.Worksheets.Add(After: lastSheet) as ExcelInterop.Worksheet;
            worksheet.Name = name;
            return worksheet;
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

        private static void ClearTableRows(ExcelInterop.Worksheet worksheet, string tableName, int columnsToClear)
        {
            var usedRange = worksheet.UsedRange;
            if (usedRange == null || usedRange.Rows.Count == 0)
            {
                return;
            }

            var startRow = usedRange.Row;
            var endRow = startRow + usedRange.Rows.Count - 1;
            var targetColumns = Math.Max(columnsToClear, 1);

            for (var row = startRow; row <= endRow; row++)
            {
                var cell = worksheet.Cells[row, 1] as ExcelInterop.Range;
                if (cell?.Value2 is not string storedTable)
                {
                    continue;
                }

                if (!string.Equals(storedTable, tableName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var range = worksheet.Range[
                    worksheet.Cells[row, 1],
                    worksheet.Cells[row, targetColumns]];
                range.ClearContents();
            }
        }

        private static int GetFirstEmptyRow(ExcelInterop.Worksheet worksheet)
        {
            var usedRange = worksheet.UsedRange;
            if (usedRange == null || usedRange.Rows.Count == 0)
            {
                return 1;
            }

            var lastRow = usedRange.Row + usedRange.Rows.Count - 1;

            for (var row = 1; row <= lastRow; row++)
            {
                var cell = worksheet.Cells[row, 1] as ExcelInterop.Range;
                if (cell?.Value2 == null)
                {
                    return row;
                }
            }

            return lastRow + 1;
        }

        private static void WriteRow(ExcelInterop.Worksheet worksheet, int row, string tableName, string[] values, int dataColumns)
        {
            var totalColumns = 1 + Math.Max(dataColumns, 0);

            for (var column = 0; column < totalColumns; column++)
            {
                var cell = worksheet.Cells[row, column + 1] as ExcelInterop.Range;

                if (column == 0)
                {
                    cell.Value2 = tableName;
                    continue;
                }

                var valueIndex = column - 1;
                var entry = valueIndex < values.Length ? values[valueIndex] : null;
                cell.Value2 = entry;
            }
        }

        private static int DetermineActualColumnCount(ExcelInterop.Worksheet worksheet, int row, int dataColumns)
        {
            for (var columnIndex = dataColumns; columnIndex >= 1; columnIndex--)
            {
                var cell = worksheet.Cells[row, columnIndex + 1] as ExcelInterop.Range;
                if (cell?.Value2 != null)
                {
                    return columnIndex;
                }
            }

            return 0;
        }
    }
}
