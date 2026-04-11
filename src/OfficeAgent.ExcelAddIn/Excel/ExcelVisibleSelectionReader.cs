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
    }
}
