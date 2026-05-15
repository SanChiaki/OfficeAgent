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

                ApplySheetNameRowFormatting(worksheet, sheetName);

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

        private static void ApplySheetNameRowFormatting(ExcelInterop.Worksheet worksheet, string sheetName)
        {
            if (worksheet == null || string.IsNullOrWhiteSpace(sheetName))
            {
                return;
            }

            var usedRange = worksheet.UsedRange;
            var rowCount = usedRange?.Rows?.Count ?? 0;
            var maxRow = Math.Min(50, Math.Max(rowCount, 50));
            for (var rowIndex = 1; rowIndex <= maxRow; rowIndex++)
            {
                var value = Convert.ToString((worksheet.Cells[rowIndex, 1] as ExcelInterop.Range)?.Value2) ?? string.Empty;
                if (!string.Equals(value, sheetName, StringComparison.Ordinal))
                {
                    continue;
                }

                FormatWorksheetRow(worksheet, rowIndex, hidden: false);
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
                    FormatWorksheetRow(worksheet, rowIndex + 1, hidden: true);
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

        private static void FormatWorksheetRow(ExcelInterop.Worksheet worksheet, int rowIndex, bool hidden)
        {
            var rowRange = worksheet.Rows[rowIndex] as ExcelInterop.Range;
            if (rowRange == null)
            {
                return;
            }

            try
            {
                rowRange.Font.Bold = true;
                rowRange.Font.Color = ColorTranslator.ToOle(Color.Blue);
                if (hidden)
                {
                    rowRange.Hidden = true;
                }
            }
            catch
            {
                // Preserve metadata writes even if formatting is not supported by the host.
            }
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
            var usedRange = worksheet.UsedRange;
            if (usedRange == null || usedRange.Rows.Count == 0 || usedRange.Columns.Count == 0)
            {
                return Array.Empty<string[]>();
            }

            var rowCount = usedRange.Rows.Count;
            var columnCount = usedRange.Columns.Count;
            var rawValues = usedRange.Value2;
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
