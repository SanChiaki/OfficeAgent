using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetColumnSegmentBuilder
    {
        public WorksheetColumnSegment[] Build(IReadOnlyList<WorksheetRuntimeColumn> columns)
        {
            var ordered = (columns ?? Array.Empty<WorksheetRuntimeColumn>())
                .Where(column => column != null)
                .OrderBy(column => column.ColumnIndex)
                .ToArray();
            if (ordered.Length == 0)
            {
                return Array.Empty<WorksheetColumnSegment>();
            }

            var deduplicated = new List<WorksheetRuntimeColumn> { ordered[0] };
            for (var index = 1; index < ordered.Length; index++)
            {
                var column = ordered[index];
                if (column.ColumnIndex == deduplicated[deduplicated.Count - 1].ColumnIndex)
                {
                    continue;
                }

                deduplicated.Add(column);
            }

            ordered = deduplicated.ToArray();

            var segments = new List<WorksheetColumnSegment>();
            var currentColumns = new List<WorksheetRuntimeColumn> { ordered[0] };
            var currentStart = ordered[0].ColumnIndex;
            var currentEnd = ordered[0].ColumnIndex;

            for (var index = 1; index < ordered.Length; index++)
            {
                var column = ordered[index];
                if (column.ColumnIndex == currentEnd + 1)
                {
                    currentColumns.Add(column);
                    currentEnd = column.ColumnIndex;
                    continue;
                }

                segments.Add(new WorksheetColumnSegment
                {
                    StartColumn = currentStart,
                    EndColumn = currentEnd,
                    Columns = currentColumns.ToArray(),
                });

                currentColumns = new List<WorksheetRuntimeColumn> { column };
                currentStart = column.ColumnIndex;
                currentEnd = column.ColumnIndex;
            }

            segments.Add(new WorksheetColumnSegment
            {
                StartColumn = currentStart,
                EndColumn = currentEnd,
                Columns = currentColumns.ToArray(),
            });

            return segments.ToArray();
        }
    }
}
