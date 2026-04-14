using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class WorksheetSchemaLayoutService
    {
        public HeaderCellPlan[] BuildHeaderPlan(WorksheetSchema schema)
        {
            if (schema == null)
            {
                throw new ArgumentNullException(nameof(schema));
            }

            var runtimeColumns = (schema.Columns ?? Array.Empty<WorksheetColumnBinding>())
                .Where(column => column != null)
                .Select(column => new WorksheetRuntimeColumn
                {
                    ColumnIndex = column.ColumnIndex,
                    ApiFieldKey = column.ApiFieldKey,
                    HeaderType = column.ColumnKind == WorksheetColumnKind.ActivityProperty ? "activityProperty" : "single",
                    DisplayText = column.ChildHeaderText,
                    ParentDisplayText = column.ParentHeaderText,
                    ChildDisplayText = column.ChildHeaderText,
                    IsIdColumn = column.IsIdColumn,
                })
                .ToArray();

            return BuildHeaderPlan(
                new SheetBinding
                {
                    HeaderStartRow = 1,
                    HeaderRowCount = 2,
                },
                runtimeColumns);
        }

        public HeaderCellPlan[] BuildHeaderPlan(SheetBinding binding, IReadOnlyList<WorksheetRuntimeColumn> columns)
        {
            if (binding == null)
            {
                throw new ArgumentNullException(nameof(binding));
            }

            var runtimeColumns = (columns ?? Array.Empty<WorksheetRuntimeColumn>())
                .Where(column => column != null)
                .OrderBy(column => column.ColumnIndex)
                .ToArray();
            var startRow = binding.HeaderStartRow <= 0 ? 1 : binding.HeaderStartRow;

            if (binding.HeaderRowCount <= 1)
            {
                return runtimeColumns
                    .Select(column => new HeaderCellPlan
                    {
                        Row = startRow,
                        Column = column.ColumnIndex,
                        Text = column.DisplayText,
                    })
                    .ToArray();
            }

            var cells = new List<HeaderCellPlan>();

            foreach (var column in runtimeColumns.Where(column => !IsActivityProperty(column)))
            {
                cells.Add(new HeaderCellPlan
                {
                    Row = startRow,
                    Column = column.ColumnIndex,
                    RowSpan = 2,
                    Text = column.DisplayText,
                });
            }

            foreach (var group in BuildActivityGroups(runtimeColumns.Where(IsActivityProperty)))
            {
                cells.Add(new HeaderCellPlan
                {
                    Row = startRow,
                    Column = group[0].ColumnIndex,
                    ColumnSpan = group.Count,
                    Text = group[0].ParentDisplayText,
                });

                foreach (var column in group)
                {
                    cells.Add(new HeaderCellPlan
                    {
                        Row = startRow + 1,
                        Column = column.ColumnIndex,
                        Text = column.ChildDisplayText,
                    });
                }
            }

            return cells
                .OrderBy(cell => cell.Row)
                .ThenBy(cell => cell.Column)
                .ToArray();
        }

        private static bool IsActivityProperty(WorksheetRuntimeColumn column)
        {
            return string.Equals(column?.HeaderType, "activityProperty", StringComparison.OrdinalIgnoreCase);
        }

        private static List<List<WorksheetRuntimeColumn>> BuildActivityGroups(IEnumerable<WorksheetRuntimeColumn> columns)
        {
            var groups = new List<List<WorksheetRuntimeColumn>>();
            List<WorksheetRuntimeColumn> currentGroup = null;
            var currentChildTexts = new HashSet<string>(StringComparer.Ordinal);

            foreach (var column in columns.OrderBy(item => item.ColumnIndex))
            {
                if (currentGroup == null ||
                    !string.Equals(currentGroup[0].ParentDisplayText, column.ParentDisplayText, StringComparison.Ordinal) ||
                    currentChildTexts.Contains(column.ChildDisplayText ?? string.Empty))
                {
                    currentGroup = new List<WorksheetRuntimeColumn>();
                    groups.Add(currentGroup);
                    currentChildTexts = new HashSet<string>(StringComparer.Ordinal);
                }

                currentGroup.Add(column);
                currentChildTexts.Add(column.ChildDisplayText ?? string.Empty);
            }

            return groups;
        }
    }
}
