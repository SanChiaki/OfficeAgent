using System;
using System.Collections.Generic;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class ExcelVisibleSelectionReader : IWorksheetSelectionReader
    {
        private readonly ExcelInterop.Application application;

        public ExcelVisibleSelectionReader(ExcelInterop.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public IReadOnlyList<SelectedVisibleCell> ReadVisibleSelection()
        {
            var selection = application.Selection as ExcelInterop.Range;
            if (selection == null)
            {
                return Array.Empty<SelectedVisibleCell>();
            }

            var results = new List<SelectedVisibleCell>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var areaCount = selection.Areas == null ? 1 : selection.Areas.Count;

            for (var areaIndex = 1; areaIndex <= areaCount; areaIndex++)
            {
                var area = selection.Areas == null
                    ? selection
                    : selection.Areas[areaIndex] as ExcelInterop.Range;

                if (area == null)
                {
                    continue;
                }

                var rowCount = Convert.ToInt32(area.Rows.Count);
                var columnCount = Convert.ToInt32(area.Columns.Count);

                for (var rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                {
                    for (var columnIndex = 1; columnIndex <= columnCount; columnIndex++)
                    {
                        var cell = area.Cells[rowIndex, columnIndex] as ExcelInterop.Range;
                        if (cell == null)
                        {
                            continue;
                        }

                        if (cell.EntireRow?.Hidden is bool rowHidden && rowHidden)
                        {
                            continue;
                        }

                        if (cell.EntireColumn?.Hidden is bool columnHidden && columnHidden)
                        {
                            continue;
                        }

                        var row = Convert.ToInt32(cell.Row);
                        var column = Convert.ToInt32(cell.Column);
                        var key = $"{row}|{column}";
                        if (!seen.Add(key))
                        {
                            continue;
                        }

                        results.Add(new SelectedVisibleCell
                        {
                            Row = row,
                            Column = column,
                            Value = Convert.ToString(cell.Text) ?? string.Empty,
                        });
                    }
                }
            }

            return results;
        }

        public WorksheetSelectionSnapshot ReadSelectionSnapshot()
        {
            var selection = application.Selection as ExcelInterop.Range;
            if (selection == null)
            {
                return new WorksheetSelectionSnapshot();
            }

            ExcelInterop.Range visibleSelection;
            try
            {
                visibleSelection = selection.SpecialCells(ExcelInterop.XlCellType.xlCellTypeVisible) as ExcelInterop.Range;
            }
            catch
            {
                visibleSelection = selection;
            }

            var areas = new List<WorksheetSelectionArea>();
            var areaSource = visibleSelection ?? selection;
            var areaCount = areaSource.Areas == null ? 1 : areaSource.Areas.Count;

            for (var areaIndex = 1; areaIndex <= areaCount; areaIndex++)
            {
                var area = areaSource.Areas == null
                    ? areaSource
                    : areaSource.Areas[areaIndex] as ExcelInterop.Range;
                if (area == null)
                {
                    continue;
                }

                var startRow = Convert.ToInt32(area.Row);
                var startColumn = Convert.ToInt32(area.Column);
                var rowCount = Convert.ToInt32(area.Rows.Count);
                var columnCount = Convert.ToInt32(area.Columns.Count);
                if (rowCount <= 0 || columnCount <= 0)
                {
                    continue;
                }

                areas.Add(new WorksheetSelectionArea
                {
                    StartRow = startRow,
                    EndRow = startRow + rowCount - 1,
                    StartColumn = startColumn,
                    EndColumn = startColumn + columnCount - 1,
                });
            }

            return new WorksheetSelectionSnapshot
            {
                Areas = areas.ToArray(),
            };
        }
    }
}
