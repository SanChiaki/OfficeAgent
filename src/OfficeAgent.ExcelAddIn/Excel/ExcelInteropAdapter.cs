using System;
using System.Collections.Generic;
using OfficeAgent.Core.Excel;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using ExcelInterop = Microsoft.Office.Interop.Excel;

namespace OfficeAgent.ExcelAddIn.Excel
{
    internal sealed class ExcelInteropAdapter : IExcelCommandExecutor
    {
        private readonly ExcelInterop.Application application;
        private readonly IExcelContextService excelContextService;

        public ExcelInteropAdapter(ExcelInterop.Application application, IExcelContextService excelContextService)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            this.excelContextService = excelContextService ?? throw new ArgumentNullException(nameof(excelContextService));
        }

        public ExcelCommandResult Preview(ExcelCommand command)
        {
            if (command == null)
            {
                throw new ArgumentNullException(nameof(command));
            }

            switch (command.CommandType)
            {
                case ExcelCommandTypes.WriteRange:
                    {
                        var target = ResolveTarget(command);
                        var targetRange = target.Worksheet.Range[target.Address];
                        ExcelOperationGuard.EnsureWorksheetAllowsMutation(target.Worksheet.Name, "write values", IsWorksheetProtected(target.Worksheet));
                        ExcelOperationGuard.EnsureRangeAllowsWrite(target.Worksheet.Name, target.Address, HasMergedCells(targetRange));

                        return new ExcelCommandResult
                        {
                            CommandType = command.CommandType,
                            RequiresConfirmation = true,
                            Status = "preview",
                            Message = $"Confirm writing values to {target.Worksheet.Name}!{target.Address}.",
                            Preview = new ExcelCommandPreview
                            {
                                Title = "Write range",
                                Summary = $"Write {command.Values.Length} row(s) x {command.Values[0].Length} column(s) to {target.Worksheet.Name}!{target.Address}",
                                Details = BuildValuePreview(command.Values),
                            },
                            SelectionContext = excelContextService.GetCurrentSelectionContext(),
                        };
                    }
                case ExcelCommandTypes.AddWorksheet:
                    ExcelOperationGuard.EnsureWorkbookStructureAllowsMutation("add worksheets", GetWorkbook().ProtectStructure);
                    return new ExcelCommandResult
                    {
                        CommandType = command.CommandType,
                        RequiresConfirmation = true,
                        Status = "preview",
                        Message = "Confirm worksheet creation before Excel is modified.",
                        Preview = new ExcelCommandPreview
                        {
                            Title = "Add worksheet",
                            Summary = $"Add worksheet \"{command.NewSheetName}\"",
                            Details = new[] { $"Workbook: {GetWorkbook().Name}" },
                        },
                        SelectionContext = excelContextService.GetCurrentSelectionContext(),
                    };
                case ExcelCommandTypes.RenameWorksheet:
                    {
                        ExcelOperationGuard.EnsureWorkbookStructureAllowsMutation("rename worksheets", GetWorkbook().ProtectStructure);
                        var worksheet = GetWorksheet(command.SheetName);
                        ExcelOperationGuard.EnsureWorksheetAllowsMutation(worksheet.Name, "rename worksheets", IsWorksheetProtected(worksheet));

                        return new ExcelCommandResult
                        {
                            CommandType = command.CommandType,
                            RequiresConfirmation = true,
                            Status = "preview",
                            Message = "Confirm worksheet rename before Excel is modified.",
                            Preview = new ExcelCommandPreview
                            {
                                Title = "Rename worksheet",
                                Summary = $"Rename worksheet \"{command.SheetName}\" to \"{command.NewSheetName}\"",
                                Details = new[] { $"Workbook: {GetWorkbook().Name}" },
                            },
                            SelectionContext = excelContextService.GetCurrentSelectionContext(),
                        };
                    }
                case ExcelCommandTypes.DeleteWorksheet:
                    {
                        ExcelOperationGuard.EnsureWorkbookStructureAllowsMutation("delete worksheets", GetWorkbook().ProtectStructure);
                        var worksheet = GetWorksheet(command.SheetName);
                        ExcelOperationGuard.EnsureWorksheetAllowsMutation(worksheet.Name, "delete worksheets", IsWorksheetProtected(worksheet));

                        return new ExcelCommandResult
                        {
                            CommandType = command.CommandType,
                            RequiresConfirmation = true,
                            Status = "preview",
                            Message = "Confirm worksheet deletion before Excel is modified.",
                            Preview = new ExcelCommandPreview
                            {
                                Title = "Delete worksheet",
                                Summary = $"Delete worksheet \"{command.SheetName}\"",
                                Details = new[] { $"Workbook: {GetWorkbook().Name}" },
                            },
                            SelectionContext = excelContextService.GetCurrentSelectionContext(),
                        };
                    }
                default:
                    throw new ArgumentException($"Excel command type '{command.CommandType}' cannot be previewed.");
            }
        }

        public ExcelCommandResult Execute(ExcelCommand command)
        {
            if (command == null)
            {
                throw new ArgumentNullException(nameof(command));
            }

            switch (command.CommandType)
            {
                case ExcelCommandTypes.ReadSelectionTable:
                    return ExecuteReadSelectionTable();
                case ExcelCommandTypes.ReadRange:
                    return ExecuteReadRange(command);
                case ExcelCommandTypes.WriteRange:
                    return ExecuteWriteRange(command);
                case ExcelCommandTypes.AddWorksheet:
                    return ExecuteAddWorksheet(command);
                case ExcelCommandTypes.RenameWorksheet:
                    return ExecuteRenameWorksheet(command);
                case ExcelCommandTypes.DeleteWorksheet:
                    return ExecuteDeleteWorksheet(command);
                default:
                    throw new ArgumentException($"Excel command type '{command.CommandType}' is not supported.");
            }
        }

        private ExcelCommandResult ExecuteReadSelectionTable()
        {
            var context = excelContextService.GetCurrentSelectionContext();
            var selection = application.Selection as ExcelInterop.Range;
            var worksheet = application.ActiveSheet as ExcelInterop.Worksheet;

            if (selection == null || worksheet == null)
            {
                throw new InvalidOperationException("No Excel range is selected.");
            }

            ExcelOperationGuard.EnsureSelectionSupportsTableRead(context, HasMergedCells(selection));

            var values = ReadRangeValues(selection);
            var headers = values.Length > 0 ? values[0] : Array.Empty<string>();
            var rows = new List<string[]>();
            for (var rowIndex = 1; rowIndex < values.Length; rowIndex++)
            {
                rows.Add(values[rowIndex]);
            }

            return new ExcelCommandResult
            {
                CommandType = ExcelCommandTypes.ReadSelectionTable,
                RequiresConfirmation = false,
                Status = "completed",
                Message = $"Read selection from {worksheet.Name} {context.Address}.",
                Table = new ExcelTableData
                {
                    SheetName = worksheet.Name,
                    Address = context.Address,
                    Headers = headers,
                    Rows = rows.ToArray(),
                },
                SelectionContext = context,
            };
        }

        private ExcelCommandResult ExecuteReadRange(ExcelCommand command)
        {
            var sheetName = command.SheetName?.Trim() ?? string.Empty;
            var address = command.TargetAddress?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(address))
            {
                throw new InvalidOperationException("A target address is required for readRange.");
            }

            ExcelInterop.Worksheet worksheet;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                worksheet = application.ActiveSheet as ExcelInterop.Worksheet;
            }
            else
            {
                worksheet = GetWorksheet(sheetName);
            }

            if (worksheet == null)
            {
                throw new InvalidOperationException("No worksheet is available for this readRange command.");
            }

            var range = worksheet.Range[address];
            if (range == null)
            {
                throw new InvalidOperationException($"Range '{address}' could not be resolved on worksheet '{worksheet.Name}'.");
            }

            var context = excelContextService.GetCurrentSelectionContext();
            var values = ReadRangeValues(range);
            var headers = values.Length > 0 ? values[0] : Array.Empty<string>();
            var rows = new List<string[]>();
            for (var rowIndex = 1; rowIndex < values.Length; rowIndex++)
            {
                rows.Add(values[rowIndex]);
            }

            return new ExcelCommandResult
            {
                CommandType = ExcelCommandTypes.ReadRange,
                RequiresConfirmation = false,
                Status = "completed",
                Message = $"Read range from {worksheet.Name} {address}.",
                Table = new ExcelTableData
                {
                    SheetName = worksheet.Name,
                    Address = address,
                    Headers = headers,
                    Rows = rows.ToArray(),
                },
                SelectionContext = context,
            };
        }

        private ExcelCommandResult ExecuteWriteRange(ExcelCommand command)
        {
            var target = ResolveTarget(command);
            var rowCount = command.Values.Length;
            var columnCount = command.Values[0].Length;
            var range = target.Worksheet.Range[target.Address];
            ExcelOperationGuard.EnsureWorksheetAllowsMutation(target.Worksheet.Name, "write values", IsWorksheetProtected(target.Worksheet));
            ExcelOperationGuard.EnsureRangeAllowsWrite(target.Worksheet.Name, target.Address, HasMergedCells(range));

            var writeTarget = range.get_Resize(rowCount, columnCount);
            writeTarget.Value2 = ToObjectArray(command.Values);

            return new ExcelCommandResult
            {
                CommandType = command.CommandType,
                RequiresConfirmation = false,
                Status = "completed",
                Message = $"Wrote {rowCount} row(s) x {columnCount} column(s) to {target.Worksheet.Name}!{target.Address}.",
                SelectionContext = excelContextService.GetCurrentSelectionContext(),
            };
        }

        private ExcelCommandResult ExecuteAddWorksheet(ExcelCommand command)
        {
            var workbook = GetWorkbook();
            ExcelOperationGuard.EnsureWorkbookStructureAllowsMutation("add worksheets", workbook.ProtectStructure);

            var worksheetCount = workbook.Worksheets.Count;
            var worksheet = workbook.Worksheets.Add(System.Type.Missing, workbook.Worksheets[worksheetCount], System.Type.Missing, System.Type.Missing) as ExcelInterop.Worksheet;
            worksheet.Name = command.NewSheetName;

            return new ExcelCommandResult
            {
                CommandType = command.CommandType,
                RequiresConfirmation = false,
                Status = "completed",
                Message = $"Worksheet \"{command.NewSheetName}\" created.",
                SelectionContext = excelContextService.GetCurrentSelectionContext(),
            };
        }

        private ExcelCommandResult ExecuteRenameWorksheet(ExcelCommand command)
        {
            ExcelOperationGuard.EnsureWorkbookStructureAllowsMutation("rename worksheets", GetWorkbook().ProtectStructure);
            var worksheet = GetWorksheet(command.SheetName);
            ExcelOperationGuard.EnsureWorksheetAllowsMutation(worksheet.Name, "rename worksheets", IsWorksheetProtected(worksheet));
            worksheet.Name = command.NewSheetName;

            return new ExcelCommandResult
            {
                CommandType = command.CommandType,
                RequiresConfirmation = false,
                Status = "completed",
                Message = $"Worksheet \"{command.SheetName}\" renamed to \"{command.NewSheetName}\".",
                SelectionContext = excelContextService.GetCurrentSelectionContext(),
            };
        }

        private ExcelCommandResult ExecuteDeleteWorksheet(ExcelCommand command)
        {
            var workbook = GetWorkbook();
            if (workbook.Worksheets.Count <= 1)
            {
                throw new InvalidOperationException("Excel must keep at least one worksheet.");
            }

            ExcelOperationGuard.EnsureWorkbookStructureAllowsMutation("delete worksheets", workbook.ProtectStructure);
            var worksheet = GetWorksheet(command.SheetName);
            ExcelOperationGuard.EnsureWorksheetAllowsMutation(worksheet.Name, "delete worksheets", IsWorksheetProtected(worksheet));

            var previousDisplayAlerts = application.DisplayAlerts;
            try
            {
                application.DisplayAlerts = false;
                worksheet.Delete();
            }
            finally
            {
                application.DisplayAlerts = previousDisplayAlerts;
            }

            return new ExcelCommandResult
            {
                CommandType = command.CommandType,
                RequiresConfirmation = false,
                Status = "completed",
                Message = $"Worksheet \"{command.SheetName}\" deleted.",
                SelectionContext = excelContextService.GetCurrentSelectionContext(),
            };
        }

        private ExcelInterop.Workbook GetWorkbook()
        {
            var workbook = application.ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException("No active workbook is available.");
            }

            return workbook;
        }

        private ExcelInterop.Worksheet GetWorksheet(string worksheetName)
        {
            var workbook = GetWorkbook();
            for (var index = 1; index <= workbook.Worksheets.Count; index++)
            {
                var worksheet = workbook.Worksheets[index] as ExcelInterop.Worksheet;
                if (worksheet != null &&
                    string.Equals(worksheet.Name, worksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return worksheet;
                }
            }

            throw new InvalidOperationException($"Worksheet '{worksheetName}' was not found in the active workbook. The workbook may have been closed or changed.");
        }

        private (ExcelInterop.Worksheet Worksheet, string Address) ResolveTarget(ExcelCommand command)
        {
            var sheetName = command.SheetName;
            var address = command.TargetAddress?.Trim() ?? string.Empty;

            if (address.Contains("!"))
            {
                var segments = address.Split(new[] { '!' }, 2);
                if (segments.Length == 2)
                {
                    if (string.IsNullOrWhiteSpace(sheetName))
                    {
                        sheetName = segments[0];
                    }
                    else if (!string.IsNullOrWhiteSpace(segments[0]) &&
                             !string.Equals(sheetName.Trim(), segments[0].Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        throw new ArgumentException("Write range commands cannot specify conflicting sheet names.");
                    }

                    address = segments[1];
                }
            }

            if (string.IsNullOrWhiteSpace(sheetName))
            {
                sheetName = (application.ActiveSheet as ExcelInterop.Worksheet)?.Name;
            }

            if (string.IsNullOrWhiteSpace(sheetName))
            {
                throw new InvalidOperationException("No worksheet is available for this Excel command.");
            }

            return (GetWorksheet(sheetName), address);
        }

        private static string[][] ReadRangeValues(ExcelInterop.Range selection)
        {
            var rowCount = Convert.ToInt32(selection.Rows.Count);
            var columnCount = Convert.ToInt32(selection.Columns.Count);
            var values = new string[rowCount][];

            for (var rowIndex = 1; rowIndex <= rowCount; rowIndex++)
            {
                var row = new string[columnCount];
                for (var columnIndex = 1; columnIndex <= columnCount; columnIndex++)
                {
                    var cell = selection.Cells[rowIndex, columnIndex] as ExcelInterop.Range;
                    row[columnIndex - 1] = Convert.ToString(cell?.Text) ?? string.Empty;
                }

                values[rowIndex - 1] = row;
            }

            return values;
        }

        private static object[,] ToObjectArray(string[][] values)
        {
            var result = new object[values.Length, values[0].Length];
            for (var rowIndex = 0; rowIndex < values.Length; rowIndex++)
            {
                for (var columnIndex = 0; columnIndex < values[rowIndex].Length; columnIndex++)
                {
                    result[rowIndex, columnIndex] = values[rowIndex][columnIndex];
                }
            }

            return result;
        }

        private static string[] BuildValuePreview(string[][] values)
        {
            var previewCount = Math.Min(values.Length, 3);
            var details = new string[previewCount];
            for (var index = 0; index < previewCount; index++)
            {
                details[index] = string.Join(" | ", values[index]);
            }

            return details;
        }

        private static bool HasMergedCells(ExcelInterop.Range range)
        {
            return range?.MergeCells is bool hasMergedCells && hasMergedCells;
        }

        private static bool IsWorksheetProtected(ExcelInterop.Worksheet worksheet)
        {
            return worksheet != null && worksheet.ProtectContents;
        }
    }
}
