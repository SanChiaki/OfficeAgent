using System;
using System.IO;
using System.Reflection;
using OfficeAgent.Core.Models;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class ProjectLayoutDialogTests
    {
        [Fact]
        public void TryCreateBindingRejectsNonNumericHeaderStartRow()
        {
            var method = GetTryCreateBindingMethod();
            var seed = CreateSeedBinding();
            var args = new object[] { seed, "abc", "2", "3", null, null };

            var success = (bool)method.Invoke(null, args);

            Assert.False(success);
            Assert.Null(args[4]);
            Assert.Equal("HeaderStartRow 必须是正整数。", (string)args[5]);
        }

        [Fact]
        public void TryCreateBindingRejectsDataStartInsideHeaderArea()
        {
            var method = GetTryCreateBindingMethod();
            var seed = CreateSeedBinding();
            var args = new object[] { seed, "1", "2", "2", null, null };

            var success = (bool)method.Invoke(null, args);

            Assert.False(success);
            Assert.Null(args[4]);
            Assert.Equal(
                "DataStartRow 必须大于或等于 HeaderStartRow + HeaderRowCount。",
                (string)args[5]);
        }

        [Fact]
        public void TryCreateBindingRejectsNonNumericHeaderRowCount()
        {
            var method = GetTryCreateBindingMethod();
            var seed = CreateSeedBinding();
            var args = new object[] { seed, "1", "abc", "3", null, null };

            var success = (bool)method.Invoke(null, args);

            Assert.False(success);
            Assert.Null(args[4]);
            Assert.Equal("HeaderRowCount 必须是正整数。", (string)args[5]);
        }

        [Fact]
        public void TryCreateBindingRejectsNonNumericDataStartRow()
        {
            var method = GetTryCreateBindingMethod();
            var seed = CreateSeedBinding();
            var args = new object[] { seed, "1", "2", "abc", null, null };

            var success = (bool)method.Invoke(null, args);

            Assert.False(success);
            Assert.Null(args[4]);
            Assert.Equal("DataStartRow 必须是正整数。", (string)args[5]);
        }

        [Fact]
        public void TryCreateBindingReturnsEditedBindingForValidValues()
        {
            var method = GetTryCreateBindingMethod();
            var seed = CreateSeedBinding();
            var args = new object[] { seed, "4", "1", "5", null, null };

            var success = (bool)method.Invoke(null, args);

            Assert.True(success);
            var binding = Assert.IsType<SheetBinding>(args[4]);
            Assert.NotSame(seed, binding);
            Assert.Equal("Sheet1", binding.SheetName);
            Assert.Equal("current-business-system", binding.SystemKey);
            Assert.Equal("performance", binding.ProjectId);
            Assert.Equal("绩效项目", binding.ProjectName);
            Assert.Equal(4, binding.HeaderStartRow);
            Assert.Equal(1, binding.HeaderRowCount);
            Assert.Equal(5, binding.DataStartRow);
            Assert.Equal(1, seed.HeaderStartRow);
            Assert.Equal(2, seed.HeaderRowCount);
            Assert.Equal(3, seed.DataStartRow);
            Assert.Null(args[5]);
        }

        private static MethodInfo GetTryCreateBindingMethod()
        {
            return Assembly.LoadFrom(ResolveAddInAssemblyPath())
                .GetType("OfficeAgent.ExcelAddIn.Dialogs.ProjectLayoutDialog", throwOnError: true)
                .GetMethod(
                    "TryCreateBinding",
                    BindingFlags.Static | BindingFlags.NonPublic)
                ?? throw new InvalidOperationException("ProjectLayoutDialog.TryCreateBinding was not found.");
        }

        private static SheetBinding CreateSeedBinding()
        {
            return new SheetBinding
            {
                SheetName = "Sheet1",
                SystemKey = "current-business-system",
                ProjectId = "performance",
                ProjectName = "绩效项目",
                HeaderStartRow = 1,
                HeaderRowCount = 2,
                DataStartRow = 3,
            };
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
