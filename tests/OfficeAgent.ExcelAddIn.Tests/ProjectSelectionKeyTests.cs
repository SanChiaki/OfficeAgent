using System;
using System.IO;
using System.Reflection;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class ProjectSelectionKeyTests
    {
        [Fact]
        public void BuildAndTryParseRoundTripSystemKeyAndProjectId()
        {
            var type = LoadType();
            var build = type.GetMethod("Build", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            var tryParse = type.GetMethod("TryParse", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);

            Assert.NotNull(build);
            Assert.NotNull(tryParse);

            var key = (string)build.Invoke(null, new object[] { "system-a", "project-1" });
            var args = new object[] { key, null, null };
            var success = (bool)tryParse.Invoke(null, args);

            Assert.True(success);
            Assert.Equal("system-a", Assert.IsType<string>(args[1]));
            Assert.Equal("project-1", Assert.IsType<string>(args[2]));
        }

        [Fact]
        public void TryParseRejectsValueWithoutSeparator()
        {
            var type = LoadType();
            var tryParse = type.GetMethod("TryParse", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);

            Assert.NotNull(tryParse);

            var args = new object[] { "project-1", null, null };
            var success = (bool)tryParse.Invoke(null, args);

            Assert.False(success);
            Assert.Null(args[1]);
            Assert.Null(args[2]);
        }

        private static Type LoadType()
        {
            return Assembly.LoadFrom(ResolveAddInAssemblyPath())
                .GetType("OfficeAgent.ExcelAddIn.ProjectSelectionKey", throwOnError: true);
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
