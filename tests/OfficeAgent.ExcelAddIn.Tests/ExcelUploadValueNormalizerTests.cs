using System;
using System.IO;
using System.Reflection;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class ExcelUploadValueNormalizerTests
    {
        [Fact]
        public void TryNormalizeReturnsPlainTextForSafeGeneralValues()
        {
            var normalizer = CreateNormalizer();

            Assert.True(InvokeTryNormalize(normalizer, null, "General", out var emptyText));
            Assert.Equal(string.Empty, emptyText);

            Assert.True(InvokeTryNormalize(normalizer, "Alpha", "@", out var stringText));
            Assert.Equal("Alpha", stringText);

            Assert.True(InvokeTryNormalize(normalizer, 12d, "General", out var integerText));
            Assert.Equal("12", integerText);

            Assert.True(InvokeTryNormalize(normalizer, 12.5d, "General", out var decimalText));
            Assert.Equal("12.5", decimalText);
        }

        [Fact]
        public void TryNormalizeReturnsFalseForDateAndFormattedNumericCells()
        {
            var normalizer = CreateNormalizer();

            Assert.False(InvokeTryNormalize(normalizer, 46024d, "yyyy/m/d", out _));
            Assert.False(InvokeTryNormalize(normalizer, 0.25d, "0%", out _));
            Assert.False(InvokeTryNormalize(normalizer, 123d, "000000", out _));
        }

        private static object CreateNormalizer()
        {
            var normalizerType = Assembly.LoadFrom(ResolveAddInAssemblyPath())
                .GetType("OfficeAgent.ExcelAddIn.Excel.ExcelUploadValueNormalizer", throwOnError: true);
            var ctor = normalizerType.GetConstructor(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                types: Type.EmptyTypes,
                modifiers: null);

            if (ctor is null)
            {
                throw new InvalidOperationException("ExcelUploadValueNormalizer constructor was not found.");
            }

            return ctor.Invoke(Array.Empty<object>());
        }

        private static bool InvokeTryNormalize(object normalizer, object value, string numberFormat, out string normalized)
        {
            var method = normalizer.GetType().GetMethod(
                "TryNormalize",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method is null)
            {
                throw new InvalidOperationException("ExcelUploadValueNormalizer.TryNormalize was not found.");
            }

            var args = new object[] { value, numberFormat, null };
            var success = (bool)method.Invoke(normalizer, args);
            normalized = args[2] as string;
            return success;
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
