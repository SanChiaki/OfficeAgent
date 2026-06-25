using System;
using System.IO;
using System.Runtime.InteropServices;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class ExcelBusinessWorkbookImporter : IBusinessWorkbookImporter
    {
        private const string BusinessDataSheetName = "Business Data";
        private readonly ExcelInterop.Application application;

        public ExcelBusinessWorkbookImporter(ExcelInterop.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public bool IsWorkSheetContentBlank(string sheetName)
        {
            var worksheet = GetWorksheet(sheetName);
            return !HasSpecialCells(worksheet, ExcelInterop.XlCellType.xlCellTypeConstants) &&
                   !HasSpecialCells(worksheet, ExcelInterop.XlCellType.xlCellTypeFormulas);
        }

        public void EnsureCanWriteToWorkSheet(string sheetName)
        {
            var workbook = GetWorkbook();
            var worksheet = GetWorksheet(sheetName);
            if (workbook.ProtectStructure)
            {
                throw new InvalidOperationException("The workbook structure is protected and cannot receive template content.");
            }

            if (worksheet.ProtectContents || worksheet.ProtectDrawingObjects || worksheet.ProtectScenarios)
            {
                throw new InvalidOperationException("The current worksheet is protected and cannot receive template content.");
            }
        }

        public void ImportBusinessDataSheet(byte[] workbookBytes, string targetSheetName)
        {
            if (workbookBytes == null || workbookBytes.Length == 0)
            {
                throw new InvalidOperationException("Business export workbook is empty.");
            }

            if (workbookBytes.Length < 2 || workbookBytes[0] != 0x50 || workbookBytes[1] != 0x4B)
            {
                throw new InvalidOperationException("Business export workbook is not a valid .xlsx file.");
            }

            EnsureCanWriteToWorkSheet(targetSheetName);
            var targetWorkbook = GetWorkbook();
            var targetWorksheet = GetWorksheet(targetWorkbook, targetSheetName);
            var originalTargetSheetName = targetWorksheet.Name;

            var tempPath = Path.Combine(
                Path.GetTempPath(),
                "OfficeAgent-BusinessExport-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelInterop.Workbook sourceWorkbook = null;
            var previousScreenUpdating = application.ScreenUpdating;
            try
            {
                application.ScreenUpdating = false;
                File.WriteAllBytes(tempPath, workbookBytes);
                sourceWorkbook = application.Workbooks.Open(
                    tempPath,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    AddToMru: false);
                TryHideWorkbookWindows(sourceWorkbook);

                var sourceWorksheet = FindWorksheet(sourceWorkbook, BusinessDataSheetName, StringComparison.Ordinal);
                if (sourceWorksheet == null)
                {
                    throw new InvalidOperationException("The exported workbook does not contain a Business Data sheet.");
                }

                CopyBusinessDataSheet(sourceWorkbook, sourceWorksheet, targetWorksheet);
                targetWorksheet.Name = originalTargetSheetName;
            }
            finally
            {
                try
                {
                    application.ScreenUpdating = previousScreenUpdating;
                }
                catch
                {
                    // Preserve the original import failure, if any.
                }

                try
                {
                    sourceWorkbook?.Close(SaveChanges: false);
                }
                catch
                {
                    // Preserve the original import failure, if any.
                }

                try
                {
                    if (File.Exists(tempPath))
                    {
                        File.Delete(tempPath);
                    }
                }
                catch
                {
                    // Temp cleanup is best-effort and must not mask import results.
                }
            }
        }

        public void ActivateWorkSheetAtA1(string sheetName)
        {
            var worksheet = GetWorksheet(sheetName);
            worksheet.Activate();
            var cell = worksheet.Range["A1"] as ExcelInterop.Range;
            cell?.Select();
        }

        private void CopyBusinessDataSheet(
            ExcelInterop.Workbook sourceWorkbook,
            ExcelInterop.Worksheet sourceWorksheet,
            ExcelInterop.Worksheet targetWorksheet)
        {
            var sourceUsedRange = sourceWorksheet.UsedRange;
            targetWorksheet.Cells.Clear();

            if (sourceUsedRange != null)
            {
                var targetStart = targetWorksheet.Range["A1"] as ExcelInterop.Range;
                sourceUsedRange.Copy(targetStart);
                TryCopyColumnWidths(sourceUsedRange, targetWorksheet);
                TryCopyRowHeights(sourceUsedRange, targetWorksheet);
            }

            TryCopyFreezePaneState(sourceWorkbook, sourceWorksheet, targetWorksheet);
        }

        private static void TryCopyColumnWidths(ExcelInterop.Range sourceUsedRange, ExcelInterop.Worksheet targetWorksheet)
        {
            try
            {
                var firstColumn = sourceUsedRange.Column;
                var columnCount = sourceUsedRange.Columns.Count;
                for (var offset = 0; offset < columnCount; offset++)
                {
                    var sourceColumn = sourceUsedRange.Worksheet.Columns[firstColumn + offset] as ExcelInterop.Range;
                    var targetColumn = targetWorksheet.Columns[1 + offset] as ExcelInterop.Range;
                    if (sourceColumn != null && targetColumn != null)
                    {
                        targetColumn.ColumnWidth = sourceColumn.ColumnWidth;
                        targetColumn.Hidden = sourceColumn.Hidden;
                    }
                }
            }
            catch
            {
                // Column presentation is best-effort; preserve the imported sheet content.
            }
        }

        private static void TryCopyRowHeights(ExcelInterop.Range sourceUsedRange, ExcelInterop.Worksheet targetWorksheet)
        {
            try
            {
                var firstRow = sourceUsedRange.Row;
                var rowCount = sourceUsedRange.Rows.Count;
                for (var offset = 0; offset < rowCount; offset++)
                {
                    var sourceRow = sourceUsedRange.Worksheet.Rows[firstRow + offset] as ExcelInterop.Range;
                    var targetRow = targetWorksheet.Rows[1 + offset] as ExcelInterop.Range;
                    if (sourceRow != null && targetRow != null)
                    {
                        targetRow.RowHeight = sourceRow.RowHeight;
                        targetRow.Hidden = sourceRow.Hidden;
                    }
                }
            }
            catch
            {
                // Row presentation is best-effort; preserve the imported sheet content.
            }
        }

        private static void TryHideWorkbookWindows(ExcelInterop.Workbook workbook)
        {
            try
            {
                if (workbook == null)
                {
                    return;
                }

                for (var index = 1; index <= workbook.Windows.Count; index++)
                {
                    var window = workbook.Windows[index] as ExcelInterop.Window;
                    if (window != null)
                    {
                        window.Visible = false;
                    }
                }
            }
            catch
            {
                // Window hiding is best-effort; import correctness is more important.
            }
        }

        private void TryCopyFreezePaneState(
            ExcelInterop.Workbook sourceWorkbook,
            ExcelInterop.Worksheet sourceWorksheet,
            ExcelInterop.Worksheet targetWorksheet)
        {
            try
            {
                var sourceWindow = sourceWorkbook.Windows.Count > 0
                    ? sourceWorkbook.Windows[1] as ExcelInterop.Window
                    : null;
                var splitRow = sourceWindow?.SplitRow ?? 0;
                var splitColumn = sourceWindow?.SplitColumn ?? 0;
                var freezePanes = sourceWindow?.FreezePanes ?? false;

                targetWorksheet.Activate();
                var targetWindow = application.ActiveWindow;
                if (targetWindow == null)
                {
                    return;
                }

                targetWindow.FreezePanes = false;
                targetWindow.SplitRow = splitRow;
                targetWindow.SplitColumn = splitColumn;
                targetWindow.FreezePanes = freezePanes;
            }
            catch
            {
                // Freeze panes are best-effort; preserve the imported sheet content.
            }
            finally
            {
                try
                {
                    targetWorksheet.Activate();
                }
                catch
                {
                    // Preserve the imported sheet content even if focus restoration fails.
                }
            }
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

        private ExcelInterop.Worksheet GetWorksheet(string sheetName)
        {
            return GetWorksheet(GetWorkbook(), sheetName);
        }

        private static ExcelInterop.Worksheet GetWorksheet(ExcelInterop.Workbook workbook, string sheetName)
        {
            var worksheet = FindWorksheet(workbook, sheetName);
            if (worksheet != null)
            {
                return worksheet;
            }

            throw new InvalidOperationException($"Worksheet '{sheetName}' was not found.");
        }

        private static ExcelInterop.Worksheet FindWorksheet(ExcelInterop.Workbook workbook, string sheetName)
        {
            return FindWorksheet(workbook, sheetName, StringComparison.OrdinalIgnoreCase);
        }

        private static ExcelInterop.Worksheet FindWorksheet(
            ExcelInterop.Workbook workbook,
            string sheetName,
            StringComparison comparison)
        {
            if (workbook == null)
            {
                return null;
            }

            for (var index = 1; index <= workbook.Worksheets.Count; index++)
            {
                var worksheet = workbook.Worksheets[index] as ExcelInterop.Worksheet;
                if (worksheet != null &&
                    string.Equals(worksheet.Name, sheetName, comparison))
                {
                    return worksheet;
                }
            }

            return null;
        }

        private static bool HasSpecialCells(ExcelInterop.Worksheet worksheet, ExcelInterop.XlCellType cellType)
        {
            try
            {
                var cells = worksheet.Cells.SpecialCells(cellType);
                return cells != null;
            }
            catch (COMException)
            {
                return false;
            }
        }
    }
}
