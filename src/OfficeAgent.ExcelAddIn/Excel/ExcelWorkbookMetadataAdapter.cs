using System;
using System.Collections.Generic;
using System.Linq;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class ExcelWorkbookMetadataAdapter : IWorksheetMetadataAdapter
    {
        private const string MetadataSheetName = "_Settings";
        private static readonly string[] OrderedTables =
        {
            "SheetBindings",
            "SheetFieldMappings",
        };

        private readonly ExcelInterop.Application application;
        private readonly MetadataSheetLayoutSerializer serializer = new MetadataSheetLayoutSerializer();

        public ExcelWorkbookMetadataAdapter(ExcelInterop.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
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

        public string[][] ReadTable(string tableName)
        {
            if (string.IsNullOrWhiteSpace(tableName))
            {
                throw new ArgumentException("Table name is required.", nameof(tableName));
            }

            return ExecutePreservingActiveWorksheet(() =>
            {
                var worksheet = EnsureWorksheetExists(MetadataSheetName);
                return serializer.ReadTable(tableName, ReadUsedRows(worksheet));
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

        private Dictionary<string, MetadataSectionDocument> LoadSections(ExcelInterop.Worksheet worksheet)
        {
            var sheetRows = ReadUsedRows(worksheet);
            var sections = new Dictionary<string, MetadataSectionDocument>(StringComparer.OrdinalIgnoreCase);

            foreach (var tableName in OrderedTables)
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
            for (var rowIndex = 0; rowIndex < rendered.Length; rowIndex++)
            {
                var values = rendered[rowIndex] ?? Array.Empty<string>();
                for (var columnIndex = 0; columnIndex < values.Length; columnIndex++)
                {
                    var cell = worksheet.Cells[rowIndex + 1, columnIndex + 1] as ExcelInterop.Range;
                    cell.Value2 = values[columnIndex];
                }
            }
        }

        private static string[][] ReadUsedRows(ExcelInterop.Worksheet worksheet)
        {
            var usedRange = worksheet.UsedRange;
            if (usedRange == null || usedRange.Rows.Count == 0 || usedRange.Columns.Count == 0)
            {
                return Array.Empty<string[]>();
            }

            var startRow = usedRange.Row;
            var rowCount = usedRange.Rows.Count;
            var columnCount = usedRange.Columns.Count;
            var rows = new string[rowCount][];

            for (var rowOffset = 0; rowOffset < rowCount; rowOffset++)
            {
                var values = new string[columnCount];
                var lastValueColumn = 0;

                for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    var cell = worksheet.Cells[startRow + rowOffset, columnIndex + 1] as ExcelInterop.Range;
                    values[columnIndex] = Convert.ToString(cell?.Value2) ?? string.Empty;
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
    }
}
