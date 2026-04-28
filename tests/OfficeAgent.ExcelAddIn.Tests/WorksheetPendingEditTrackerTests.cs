using System;
using System.IO;
using System.Reflection;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WorksheetPendingEditTrackerTests
    {
        [Fact]
        public void MarkChangedStoresCapturedOriginalValueUntilCleared()
        {
            var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var trackerType = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetPendingEditTracker", throwOnError: true);
            var tracker = Activator.CreateInstance(trackerType);

            CaptureBeforeValues(assembly, tracker, "Sheet1", (6, 4, "2026-01-05"));
            MarkChanged(assembly, tracker, "Sheet1", (6, 4));

            Assert.True(TryGetOriginalValue(tracker, "Sheet1", 6, 4, out var value));
            Assert.Equal("2026-01-05", value);

            trackerType.GetMethod("Clear", new[] { typeof(string), typeof(int), typeof(int) })
                .Invoke(tracker, new object[] { "Sheet1", 6, 4 });

            Assert.False(TryGetOriginalValue(tracker, "Sheet1", 6, 4, out _));
        }

        [Fact]
        public void MarkChangedKeepsFirstOriginalValueAcrossRepeatedEdits()
        {
            var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var trackerType = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetPendingEditTracker", throwOnError: true);
            var tracker = Activator.CreateInstance(trackerType);

            CaptureBeforeValues(assembly, tracker, "Sheet1", (6, 4, "原始值"));
            MarkChanged(assembly, tracker, "Sheet1", (6, 4));
            CaptureBeforeValues(assembly, tracker, "Sheet1", (6, 4, "中间值"));
            MarkChanged(assembly, tracker, "Sheet1", (6, 4));

            Assert.True(TryGetOriginalValue(tracker, "Sheet1", 6, 4, out var value));
            Assert.Equal("原始值", value);
        }

        private static void CaptureBeforeValues(
            Assembly assembly,
            object tracker,
            string sheetName,
            params (int Row, int Column, string Text)[] cells)
        {
            var valueType = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetCellValue", throwOnError: true);
            var values = Array.CreateInstance(valueType, cells.Length);
            for (var index = 0; index < cells.Length; index++)
            {
                var value = Activator.CreateInstance(valueType);
                SetProperty(value, "Row", cells[index].Row);
                SetProperty(value, "Column", cells[index].Column);
                SetProperty(value, "Text", cells[index].Text);
                values.SetValue(value, index);
            }

            tracker.GetType().GetMethod("CaptureBeforeValues").Invoke(tracker, new object[] { sheetName, values });
        }

        private static void MarkChanged(
            Assembly assembly,
            object tracker,
            string sheetName,
            params (int Row, int Column)[] cells)
        {
            var addressType = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetCellAddress", throwOnError: true);
            var values = Array.CreateInstance(addressType, cells.Length);
            for (var index = 0; index < cells.Length; index++)
            {
                var value = Activator.CreateInstance(addressType);
                SetProperty(value, "Row", cells[index].Row);
                SetProperty(value, "Column", cells[index].Column);
                values.SetValue(value, index);
            }

            tracker.GetType().GetMethod("MarkChanged").Invoke(tracker, new object[] { sheetName, values });
        }

        private static bool TryGetOriginalValue(object tracker, string sheetName, int row, int column, out string value)
        {
            var args = new object[] { sheetName, row, column, null };
            var result = (bool)tracker.GetType().GetMethod("TryGetOriginalValue").Invoke(tracker, args);
            value = Convert.ToString(args[3]) ?? string.Empty;
            return result;
        }

        private static void SetProperty(object target, string propertyName, object value)
        {
            target.GetType()
                .GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                .SetValue(target, value);
        }

        private static string ResolveAddInAssemblyPath()
        {
            return Path.GetFullPath(
                Path.Combine(
                    AppContext.BaseDirectory,
                    "..",
                    "..",
                    "..",
                    "..",
                    "..",
                    "src",
                    "OfficeAgent.ExcelAddIn",
                    "bin",
                    "Debug",
                    "OfficeAgent.ExcelAddIn.dll"));
        }
    }
}
