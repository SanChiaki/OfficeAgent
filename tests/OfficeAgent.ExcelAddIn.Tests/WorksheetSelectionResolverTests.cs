using System;
using System.IO;
using System.Reflection;
using OfficeAgent.Core.Models;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WorksheetSelectionResolverTests
    {
        [Fact]
        public void ResolveReturnsIdsAndFieldKeysWhenHeadersAndIdCellsAreNotSelected()
        {
            var resolver = CreateResolver();
            var schema = new WorksheetSchema
            {
                Columns = new[]
                {
                    new WorksheetColumnBinding { ColumnIndex = 1, ApiFieldKey = "id", IsIdColumn = true },
                    new WorksheetColumnBinding { ColumnIndex = 2, ApiFieldKey = "name" },
                    new WorksheetColumnBinding { ColumnIndex = 3, ApiFieldKey = "start_12345678" },
                },
            };

            var visibleCells = new[]
            {
                new SelectedVisibleCell { Row = 3, Column = 2, Value = "椤圭洰A" },
                new SelectedVisibleCell { Row = 3, Column = 3, Value = "2026-01-02" },
            };

            Func<int, string> rowIdAccessor = row => row == 3 ? "row-1" : string.Empty;

            var resolved = InvokeResolve(resolver, schema, visibleCells, rowIdAccessor);

            Assert.Equal(new[] { "row-1" }, resolved.RowIds);
            Assert.Equal(new[] { "name", "start_12345678" }, resolved.ApiFieldKeys);
            Assert.Equal(2, resolved.TargetCells.Length);
        }

        [Fact]
        public void ResolveExcludesCellsMissingSchemaBindingFromTargetCells()
        {
            var resolver = CreateResolver();
            var schema = new WorksheetSchema
            {
                Columns = new[]
                {
                    new WorksheetColumnBinding { ColumnIndex = 1, ApiFieldKey = "id", IsIdColumn = true },
                    new WorksheetColumnBinding { ColumnIndex = 2, ApiFieldKey = "name" },
                },
            };

            var visibleCells = new[]
            {
                new SelectedVisibleCell { Row = 3, Column = 2, Value = "椤圭洰A" },
                new SelectedVisibleCell { Row = 4, Column = 99, Value = "Unknown column" },
            };

            Func<int, string> rowIdAccessor = row => row == 3 ? "row-1" : "row-2";

            var resolved = InvokeResolve(resolver, schema, visibleCells, rowIdAccessor);

            Assert.DoesNotContain(resolved.TargetCells, cell => cell.Column == 99);
            Assert.Equal(new[] { "row-1" }, resolved.RowIds);
        }

        [Fact]
        public void ResolveExcludesCellsWithBlankRowIdsFromTargetCells()
        {
            var resolver = CreateResolver();
            var schema = new WorksheetSchema
            {
                Columns = new[]
                {
                    new WorksheetColumnBinding { ColumnIndex = 2, ApiFieldKey = "name" },
                    new WorksheetColumnBinding { ColumnIndex = 3, ApiFieldKey = "start_12345678" },
                },
            };

            var visibleCells = new[]
            {
                new SelectedVisibleCell { Row = 3, Column = 2, Value = "椤圭洰A" },
                new SelectedVisibleCell { Row = 4, Column = 3, Value = "2026-01-02" },
            };

            Func<int, string> rowIdAccessor = row => row == 3 ? "row-1" : string.Empty;

            var resolved = InvokeResolve(resolver, schema, visibleCells, rowIdAccessor);

            Assert.DoesNotContain(resolved.TargetCells, cell => cell.Row == 4);
            Assert.Equal(new[] { "row-1" }, resolved.RowIds);
        }

        [Fact]
        public void ResolveDoesNotIncludeIdColumnInFieldKeys()
        {
            var resolver = CreateResolver();
            var schema = new WorksheetSchema
            {
                Columns = new[]
                {
                    new WorksheetColumnBinding { ColumnIndex = 1, ApiFieldKey = "id", IsIdColumn = true },
                    new WorksheetColumnBinding { ColumnIndex = 2, ApiFieldKey = "name" },
                    new WorksheetColumnBinding { ColumnIndex = 3, ApiFieldKey = "start_12345678" },
                },
            };

            var visibleCells = new[]
            {
                new SelectedVisibleCell { Row = 3, Column = 1, Value = "row-1" },
                new SelectedVisibleCell { Row = 3, Column = 2, Value = "椤圭洰A" },
                new SelectedVisibleCell { Row = 3, Column = 3, Value = "2026-01-02" },
            };

            Func<int, string> rowIdAccessor = row => row == 3 ? "row-1" : string.Empty;

            var resolved = InvokeResolve(resolver, schema, visibleCells, rowIdAccessor);

            Assert.DoesNotContain("id", resolved.ApiFieldKeys);
            Assert.Equal(new[] { "name", "start_12345678" }, resolved.ApiFieldKeys);
        }

        private static object CreateResolver()
        {
            var assembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var resolverType = assembly.GetType("OfficeAgent.ExcelAddIn.Excel.WorksheetSelectionResolver", throwOnError: true);
            var ctor = resolverType.GetConstructor(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                types: Type.EmptyTypes,
                modifiers: null);

            if (ctor is null)
            {
                throw new InvalidOperationException("WorksheetSelectionResolver constructor was not found.");
            }

            return ctor.Invoke(Array.Empty<object>());
        }

        private static ResolvedSelection InvokeResolve(
            object resolver,
            WorksheetSchema schema,
            SelectedVisibleCell[] visibleCells,
            Func<int, string> rowIdAccessor)
        {
            var method = resolver.GetType().GetMethod(
                "Resolve",
                BindingFlags.Instance | BindingFlags.Public);

            return (ResolvedSelection)method.Invoke(resolver, new object[] { schema, visibleCells, rowIdAccessor });
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
