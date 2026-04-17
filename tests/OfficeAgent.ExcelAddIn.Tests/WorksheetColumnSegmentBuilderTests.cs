using System;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeAgent.Core.Models;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WorksheetColumnSegmentBuilderTests
    {
        [Fact]
        public void BuildGroupsContiguousManagedColumnsIntoSegments()
        {
            var builder = CreateBuilder();
            var segments = Build(builder, new[]
            {
                new WorksheetRuntimeColumn { ColumnIndex = 1, ApiFieldKey = "row_id", IsIdColumn = true },
                new WorksheetRuntimeColumn { ColumnIndex = 2, ApiFieldKey = "owner_name" },
                new WorksheetRuntimeColumn { ColumnIndex = 4, ApiFieldKey = "start_12345678" },
                new WorksheetRuntimeColumn { ColumnIndex = 5, ApiFieldKey = "end_12345678" },
            });

            Assert.Collection(
                segments,
                first =>
                {
                    Assert.Equal(1, GetStartColumn(first));
                    Assert.Equal(2, GetEndColumn(first));
                    Assert.Equal(new[] { 1, 2 }, GetColumns(first).Select(column => column.ColumnIndex).ToArray());
                },
                second =>
                {
                    Assert.Equal(4, GetStartColumn(second));
                    Assert.Equal(5, GetEndColumn(second));
                    Assert.Equal(new[] { 4, 5 }, GetColumns(second).Select(column => column.ColumnIndex).ToArray());
                });
        }

        [Fact]
        public void BuildSkipsNullColumnsAndReturnsEmptyForNoManagedColumns()
        {
            var builder = CreateBuilder();
            var segments = Build(builder, new WorksheetRuntimeColumn[] { null });

            Assert.Empty(segments);
        }

        [Fact]
        public void BuildDeduplicatesDuplicateColumnIndexesKeepingFirstAfterOrdering()
        {
            var builder = CreateBuilder();
            var firstColumnTwo = new WorksheetRuntimeColumn { ColumnIndex = 2, ApiFieldKey = "owner_name" };
            var secondColumnTwo = new WorksheetRuntimeColumn { ColumnIndex = 2, ApiFieldKey = "owner_alias" };
            var segments = Build(builder, new[]
            {
                new WorksheetRuntimeColumn { ColumnIndex = 3, ApiFieldKey = "end_12345678" },
                firstColumnTwo,
                new WorksheetRuntimeColumn { ColumnIndex = 1, ApiFieldKey = "row_id", IsIdColumn = true },
                secondColumnTwo,
            });

            Assert.Single(segments);
            var segment = segments[0];
            Assert.Equal(1, GetStartColumn(segment));
            Assert.Equal(3, GetEndColumn(segment));

            var columns = GetColumns(segment);
            Assert.Equal(new[] { 1, 2, 3 }, columns.Select(column => column.ColumnIndex).ToArray());
            Assert.Same(firstColumnTwo, columns[1]);
        }

        private static object CreateBuilder()
        {
            var builderType = LoadAddInAssembly()
                .GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetColumnSegmentBuilder", throwOnError: true);
            return Activator.CreateInstance(builderType, nonPublic: true);
        }

        private static object[] Build(object builder, WorksheetRuntimeColumn[] columns)
        {
            var method = builder.GetType().GetMethod("Build", BindingFlags.Instance | BindingFlags.Public);
            var segments = method.Invoke(builder, new object[] { columns });
            return ((Array)segments).Cast<object>().ToArray();
        }

        private static int GetStartColumn(object segment)
        {
            return (int)segment.GetType().GetProperty("StartColumn").GetValue(segment);
        }

        private static int GetEndColumn(object segment)
        {
            return (int)segment.GetType().GetProperty("EndColumn").GetValue(segment);
        }

        private static WorksheetRuntimeColumn[] GetColumns(object segment)
        {
            return (WorksheetRuntimeColumn[])segment.GetType().GetProperty("Columns").GetValue(segment);
        }

        private static Assembly LoadAddInAssembly()
        {
            return Assembly.LoadFrom(ResolveAddInAssemblyPath());
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
