using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Drawing;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class ExcelWorkbookMetadataAdapter : IWorksheetMetadataAdapter
    {
        private const string MetadataSheetName = MetadataWorksheetNames.Current;
        private const int SheetNamePresentationScanLimit = 50;

        private readonly ExcelInterop.Application application;
        private readonly MetadataSheetLayoutSerializer serializer = new MetadataSheetLayoutSerializer();

        public ExcelWorkbookMetadataAdapter(ExcelInterop.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public string GetWorkbookScopeKey()
        {
            var workbook = application.ActiveWorkbook;
            if (workbook == null)
            {
                return string.Empty;
            }

            if (!string.IsNullOrWhiteSpace(workbook.FullName))
            {
                return workbook.FullName;
            }

            if (!string.IsNullOrWhiteSpace(workbook.Name))
            {
                return workbook.Name;
            }

            return workbook.GetHashCode().ToString(CultureInfo.InvariantCulture);
        }

        public void EnsureWorksheet(string name, bool visible)
        {
            ExecutePreservingActiveWorksheet(() =>
            {
                var worksheet = EnsureWorksheetExists(name);
                worksheet.Visible = visible
                    ? ExcelInterop.XlSheetVisibility.xlSheetVisible
                    : ExcelInterop.XlSheetVisibility.xlSheetHidden;
            });
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

            ExecutePreservingActiveWorksheet(() =>
            {
                var worksheet = EnsureWorksheetExists(MetadataSheetName);
                var sections = LoadSections(worksheet);
                sections[tableName] = new MetadataSectionDocument(tableName, headers, rows);
                RewriteSheet(worksheet, sections);
            });
        }

        public void ApplyMetadataPresentation(string sheetName, bool hideTemplateBindingRows)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return;
            }

            ExecutePreservingActiveWorksheet(() =>
            {
                var worksheet = FindMetadataWorksheet(GetWorkbook());
                if (worksheet == null)
                {
                    return;
                }

                ApplySheetNameRowFormatting(worksheet);

                if (hideTemplateBindingRows)
                {
                    HideTemplateBindingRow(worksheet, sheetName);
                }
            });
        }

        public string[][] ReadTable(string tableName)
        {
            if (string.IsNullOrWhiteSpace(tableName))
            {
                throw new ArgumentException("Table name is required.", nameof(tableName));
            }

            return ExecutePreservingActiveWorksheet(() =>
            {
                var worksheet = FindMetadataWorksheet(GetWorkbook());
                if (worksheet == null)
                {
                    return Array.Empty<string[]>();
                }

                return serializer.ReadTable(tableName, ReadUsedRows(worksheet));
            });
        }

        public string[] ReadHeaders(string tableName)
        {
            if (string.IsNullOrWhiteSpace(tableName))
            {
                throw new ArgumentException("Table name is required.", nameof(tableName));
            }

            return ExecutePreservingActiveWorksheet(() =>
            {
                var worksheet = FindMetadataWorksheet(GetWorkbook());
                if (worksheet == null)
                {
                    return Array.Empty<string>();
                }

                var section = serializer.ReadSection(tableName, ReadUsedRows(worksheet));
                return section?.Headers ?? Array.Empty<string>();
            });
        }

        private void ExecutePreservingActiveWorksheet(Action action)
        {
            ExecutePreservingActiveWorksheet(() =>
            {
                action();
                return true;
            });
        }

        private T ExecutePreservingActiveWorksheet<T>(Func<T> action)
        {
            var activeSheet = application.ActiveSheet as ExcelInterop.Worksheet;

            try
            {
                return action();
            }
            finally
            {
                if (activeSheet != null)
                {
                    try
                    {
                        activeSheet.Activate();
                    }
                    catch
                    {
                        // Ignore focus restoration failures and keep metadata operations successful.
                    }
                }
            }
        }

        private ExcelInterop.Worksheet EnsureWorksheetExists(string name)
        {
            var workbook = GetWorkbook();
            if (string.Equals(name, MetadataSheetName, StringComparison.OrdinalIgnoreCase))
            {
                var metadataWorksheet = FindMetadataWorksheet(workbook);
                if (metadataWorksheet != null)
                {
                    return metadataWorksheet;
                }
            }

            var existing = FindWorksheet(workbook, name);
            if (existing != null)
            {
                return existing;
            }

            var lastSheet = workbook.Worksheets[workbook.Worksheets.Count] as ExcelInterop.Worksheet;
            var worksheet = workbook.Worksheets.Add(After: lastSheet) as ExcelInterop.Worksheet;
            worksheet.Name = name;
            return worksheet;
        }

        private static void ApplySheetNameRowFormatting(ExcelInterop.Worksheet worksheet)
        {
            if (worksheet == null)
            {
                return;
            }

            var rows = ReadUsedRows(worksheet);
            var maxRow = Math.Min(SheetNamePresentationScanLimit, Math.Max(rows.Length, SheetNamePresentationScanLimit));
            var presentationColumnCount = GetPresentationColumnCount(rows, maxRow);
            ClearSheetNameRowFormatting(worksheet, maxRow, presentationColumnCount);
            for (var rowIndex = 1; rowIndex <= maxRow; rowIndex++)
            {
                var row = rowIndex <= rows.Length
                    ? rows[rowIndex - 1] ?? Array.Empty<string>()
                    : Array.Empty<string>();
                var value = row.Length > 0 ? row[0] : string.Empty;
                if (!string.Equals(value, "SheetName", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                FormatWorksheetRow(worksheet, rowIndex, presentationColumnCount, hidden: false);
            }
        }

        private static void ClearSheetNameRowFormatting(ExcelInterop.Worksheet worksheet, int maxRow, int presentationColumnCount)
        {
            for (var rowIndex = 1; rowIndex <= maxRow; rowIndex++)
            {
                ResetWorksheetRowFormatting(worksheet, rowIndex, presentationColumnCount);
            }
        }

        private static void HideTemplateBindingRow(ExcelInterop.Worksheet worksheet, string sheetName)
        {
            if (worksheet == null || string.IsNullOrWhiteSpace(sheetName))
            {
                return;
            }

            var rows = ReadUsedRows(worksheet);
            var templateBindingsStart = Array.FindIndex(rows, row =>
                row.Length > 0 &&
                string.Equals(row[0], "TemplateBindings", StringComparison.OrdinalIgnoreCase));
            if (templateBindingsStart < 0)
            {
                return;
            }

            for (var rowIndex = templateBindingsStart + 2; rowIndex < rows.Length; rowIndex++)
            {
                var candidate = rows[rowIndex] ?? Array.Empty<string>();
                if (candidate.Length > 0 &&
                    !string.IsNullOrWhiteSpace(candidate[0]) &&
                    string.Equals(candidate[0], sheetName, StringComparison.Ordinal))
                {
                    FormatWorksheetRow(worksheet, rowIndex + 1, Math.Max(1, candidate.Length), hidden: true);
                    return;
                }

                if (candidate.Length > 0 &&
                    !string.IsNullOrWhiteSpace(candidate[0]) &&
                    Array.IndexOf(MetadataSheetLayoutSerializer.OrderedSectionNames.ToArray(), candidate[0]) >= 0)
                {
                    return;
                }
            }
        }

        private static void FormatWorksheetRow(ExcelInterop.Worksheet worksheet, int rowIndex, int presentationColumnCount, bool hidden)
        {
            var formatRange = GetPresentationRange(worksheet, rowIndex, presentationColumnCount);
            if (formatRange == null)
            {
                return;
            }

            try
            {
                formatRange.Font.Bold = true;
                formatRange.Font.Color = ColorTranslator.ToOle(Color.Blue);
                if (hidden)
                {
                    var rowRange = worksheet.Rows[rowIndex] as ExcelInterop.Range;
                    if (rowRange == null)
                    {
                        return;
                    }

                    rowRange.Hidden = true;
                }
            }
            catch
            {
                // Preserve metadata writes even if formatting is not supported by the host.
            }
        }

        private static void ResetWorksheetRowFormatting(ExcelInterop.Worksheet worksheet, int rowIndex, int presentationColumnCount)
        {
            var formatRange = GetPresentationRange(worksheet, rowIndex, presentationColumnCount);
            if (formatRange == null)
            {
                return;
            }

            try
            {
                formatRange.Font.Bold = false;
                formatRange.Font.ColorIndex = ExcelInterop.XlColorIndex.xlColorIndexAutomatic;
            }
            catch
            {
                // Preserve metadata writes even if formatting reset is not supported by the host.
            }
        }

        private static int GetPresentationColumnCount(IReadOnlyList<string[]> rows, int maxRow)
        {
            var count = 1;
            if (rows == null)
            {
                return count;
            }

            for (var rowIndex = 0; rowIndex < rows.Count && rowIndex < maxRow; rowIndex++)
            {
                count = Math.Max(count, rows[rowIndex]?.Length ?? 0);
            }

            return count;
        }

        private static ExcelInterop.Range GetPresentationRange(
            ExcelInterop.Worksheet worksheet,
            int rowIndex,
            int presentationColumnCount)
        {
            if (worksheet == null || rowIndex <= 0 || presentationColumnCount <= 0)
            {
                return null;
            }

            var startCell = worksheet.Cells[rowIndex, 1] as ExcelInterop.Range;
            return startCell?.Resize[1, presentationColumnCount] as ExcelInterop.Range;
        }

        private ExcelInterop.Worksheet FindWorksheet(string name)
        {
            return FindWorksheet(GetWorkbook(), name);
        }

        private static ExcelInterop.Worksheet FindWorksheet(ExcelInterop.Workbook workbook, string name)
        {
            foreach (ExcelInterop.Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }

            return null;
        }

        private static ExcelInterop.Worksheet FindMetadataWorksheet(ExcelInterop.Workbook workbook)
        {
            var current = FindWorksheet(workbook, MetadataWorksheetNames.Current);
            if (current != null)
            {
                return current;
            }

            var legacy = FindWorksheet(workbook, MetadataWorksheetNames.Legacy);
            if (legacy == null)
            {
                return null;
            }

            legacy.Name = MetadataWorksheetNames.Current;
            return legacy;
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

        private Dictionary<string, MetadataSectionDocument> LoadSections(ExcelInterop.Worksheet worksheet)
        {
            var sheetRows = ReadUsedRows(worksheet);
            var sections = new Dictionary<string, MetadataSectionDocument>(StringComparer.OrdinalIgnoreCase);

            foreach (var tableName in MetadataSheetLayoutSerializer.OrderedSectionNames)
            {
                var section = serializer.ReadSection(tableName, sheetRows);
                if (section == null || section.Headers.Length == 0)
                {
                    continue;
                }

                sections[tableName] = section;
            }

            return sections;
        }

        private void RewriteSheet(
            ExcelInterop.Worksheet worksheet,
            IReadOnlyDictionary<string, MetadataSectionDocument> sections)
        {
            var cells = worksheet.Cells as ExcelInterop.Range;
            cells?.ClearContents();

            var rendered = serializer.Render(sections);
            if (rendered.Length == 0)
            {
                return;
            }

            var objectValues = ToObjectMatrix(rendered, out var columnCount);
            if (columnCount <= 0)
            {
                return;
            }

            var startCell = worksheet.Cells[1, 1] as ExcelInterop.Range;
            var writeTarget = startCell?.Resize[rendered.Length, columnCount] as ExcelInterop.Range;
            if (writeTarget == null)
            {
                return;
            }

            writeTarget.Value2 = objectValues;
        }

        private static string[][] ReadUsedRows(ExcelInterop.Worksheet worksheet)
        {
            var readRange = GetActualUsedRange(worksheet) ?? worksheet.UsedRange;
            if (readRange == null || readRange.Rows.Count == 0 || readRange.Columns.Count == 0)
            {
                return Array.Empty<string[]>();
            }

            var rowCount = readRange.Rows.Count;
            var columnCount = readRange.Columns.Count;
            var rawValues = readRange.Value2;
            var rows = new string[rowCount][];

            for (var rowOffset = 0; rowOffset < rowCount; rowOffset++)
            {
                var values = new string[columnCount];
                var lastValueColumn = 0;

                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    values[columnIndex] = Convert.ToString(
                        GetRangeValue(rawValues, rowOffset, columnIndex, rowCount, columnCount)) ?? string.Empty;
                    if (!string.IsNullOrEmpty(values[columnIndex]))
                    {
                        lastValueColumn = columnIndex + 1;
                    }
                }

                rows[rowOffset] = lastValueColumn == 0
                    ? Array.Empty<string>()
                    : values.Take(lastValueColumn).ToArray();
            }

            return rows;
        }

        private static ExcelInterop.Range GetActualUsedRange(ExcelInterop.Worksheet worksheet)
        {
            try
            {
                var cells = worksheet?.Cells as ExcelInterop.Range;
                if (cells == null)
                {
                    return null;
                }

                var anchor = cells[1, 1] as ExcelInterop.Range;
                var lastRowCell = cells.Find(
                    What: "*",
                    After: anchor,
                    LookIn: ExcelInterop.XlFindLookIn.xlFormulas,
                    LookAt: ExcelInterop.XlLookAt.xlPart,
                    SearchOrder: ExcelInterop.XlSearchOrder.xlByRows,
                    SearchDirection: ExcelInterop.XlSearchDirection.xlPrevious,
                    MatchCase: false,
                    SearchFormat: false) as ExcelInterop.Range;
                var lastColumnCell = cells.Find(
                    What: "*",
                    After: anchor,
                    LookIn: ExcelInterop.XlFindLookIn.xlFormulas,
                    LookAt: ExcelInterop.XlLookAt.xlPart,
                    SearchOrder: ExcelInterop.XlSearchOrder.xlByColumns,
                    SearchDirection: ExcelInterop.XlSearchDirection.xlPrevious,
                    MatchCase: false,
                    SearchFormat: false) as ExcelInterop.Range;
                if (lastRowCell == null || lastColumnCell == null)
                {
                    return null;
                }

                var topLeft = worksheet.Cells[1, 1] as ExcelInterop.Range;
                var bottomRight = worksheet.Cells[lastRowCell.Row, lastColumnCell.Column] as ExcelInterop.Range;
                return worksheet.Range[topLeft, bottomRight] as ExcelInterop.Range;
            }
            catch
            {
                // Fall back to UsedRange when the host or tests do not support Find-based bounds.
                return null;
            }
        }

        private static object GetRangeValue(object rawValues, int rowOffset, int columnOffset, int rowCount, int columnCount)
        {
            if (!(rawValues is Array array))
            {
                return rowCount == 1 && columnCount == 1 && rowOffset == 0 && columnOffset == 0
                    ? rawValues
                    : null;
            }

            if (array.Rank != 2)
            {
                return null;
            }

            var rowIndex = array.GetLowerBound(0) + rowOffset;
            var columnIndex = array.GetLowerBound(1) + columnOffset;
            if (rowIndex > array.GetUpperBound(0) || columnIndex > array.GetUpperBound(1))
            {
                return null;
            }

            return array.GetValue(rowIndex, columnIndex);
        }

        private static object[,] ToObjectMatrix(string[][] rows, out int columnCount)
        {
            columnCount = 0;
            if (rows == null || rows.Length == 0)
            {
                return new object[0, 0];
            }

            for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
            {
                columnCount = Math.Max(columnCount, rows[rowIndex]?.Length ?? 0);
            }

            if (columnCount == 0)
            {
                return new object[rows.Length, 0];
            }

            var values = new object[rows.Length, columnCount];
            for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
            {
                var row = rows[rowIndex] ?? Array.Empty<string>();
                for (var columnIndex = 0; columnIndex < row.Length; columnIndex++)
                {
                    values[rowIndex, columnIndex] = row[columnIndex];
                }
            }

            return values;
        }
    }
}
