using System;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class ExcelSelectionContextService : IExcelContextService
    {
        private const int PreviewRowCount = 4;
        private const int PreviewColumnCount = 5;

        private readonly ExcelInterop.Application application;

        public ExcelSelectionContextService(ExcelInterop.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public SelectionContext GetCurrentSelectionContext()
        {
            try
            {
                var selection = application.Selection as ExcelInterop.Range;
                var activeWorkbook = application.ActiveWorkbook;
                var activeWorksheet = application.ActiveSheet as ExcelInterop.Worksheet;

                if (selection == null || activeWorksheet == null)
                {
                    return SelectionContext.Empty("No selection available.");
                }

                var rowCount = Convert.ToInt32(selection.Rows.Count);
                var columnCount = Convert.ToInt32(selection.Columns.Count);
                var areaCount = selection.Areas == null ? 1 : Convert.ToInt32(selection.Areas.Count);
                var address = Convert.ToString(selection.get_Address(false, false, ExcelInterop.XlReferenceStyle.xlA1));

                string[,] previewValues = null;
                if (areaCount <= 1)
                {
                    previewValues = ReadPreviewValues(selection, rowCount, columnCount);
                }

                return SelectionContextFactory.Create(
                    workbookName: activeWorkbook?.Name ?? string.Empty,
                    sheetName: activeWorksheet.Name,
                    address: address ?? string.Empty,
                    rowCount: rowCount,
                    columnCount: columnCount,
                    areaCount: areaCount,
                    previewValues: previewValues);
            }
            catch (Exception)
            {
                return SelectionContext.Empty("Unable to read the current selection.");
            }
        }

        private static string[,] ReadPreviewValues(ExcelInterop.Range selection, int rowCount, int columnCount)
        {
            var previewRowCount = Math.Min(rowCount, PreviewRowCount);
            var previewColumnCount = Math.Min(columnCount, PreviewColumnCount);
            var values = new string[previewRowCount, previewColumnCount];

            for (var rowIndex = 1; rowIndex <= previewRowCount; rowIndex++)
            {
                for (var columnIndex = 1; columnIndex <= previewColumnCount; columnIndex++)
                {
                    var cell = selection.Cells[rowIndex, columnIndex] as ExcelInterop.Range;
                    values[rowIndex - 1, columnIndex - 1] = Convert.ToString(cell?.Text) ?? string.Empty;
                }
            }

            return values;
        }
    }
}
