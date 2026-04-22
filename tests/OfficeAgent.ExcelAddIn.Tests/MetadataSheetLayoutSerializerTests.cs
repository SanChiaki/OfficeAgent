using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class MetadataSheetLayoutSerializerTests
    {
        [Fact]
        public void RenderPlacesTemplateBindingsAboveBindingsAndFieldMappingsUsingReadableSections()
        {
            var serializer = CreateSerializer();
            var rendered = InvokeRender(
                serializer,
                new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["TemplateBindings"] = CreateSection(
                        "TemplateBindings",
                        new[] { "SheetName", "TemplateId", "TemplateName" },
                        new[] { new[] { "Sheet1", "tmpl-001", "Quarterly Report" } }),
                    ["SheetBindings"] = CreateSection(
                        "SheetBindings",
                        new[] { "SheetName", "SystemKey" },
                        new[] { new[] { "Sheet1", "current-business-system" } }),
                    ["SheetFieldMappings"] = CreateSection(
                        "SheetFieldMappings",
                        new[] { "SheetName", "HeaderId", "ApiFieldKey" },
                        new[] { new[] { "Sheet1", "row_id", "row_id" } }),
                });

            Assert.Equal("TemplateBindings", rendered[0][0]);
            Assert.Equal(new[] { "SheetName", "TemplateId", "TemplateName" }, rendered[1]);
            Assert.Equal(new[] { "Sheet1", "tmpl-001", "Quarterly Report" }, rendered[2]);
            Assert.True(rendered[3].All(string.IsNullOrEmpty));
            Assert.True(rendered[4].All(string.IsNullOrEmpty));
            Assert.Equal("SheetBindings", rendered[5][0]);
            Assert.Equal(new[] { "SheetName", "SystemKey" }, rendered[6]);
            Assert.Equal(new[] { "Sheet1", "current-business-system" }, rendered[7]);
            Assert.True(rendered[8].All(string.IsNullOrEmpty));
            Assert.True(rendered[9].All(string.IsNullOrEmpty));
            Assert.Equal("SheetFieldMappings", rendered[10][0]);
            Assert.Equal(new[] { "SheetName", "HeaderId", "ApiFieldKey" }, rendered[11]);
            Assert.Equal(new[] { "Sheet1", "row_id", "row_id" }, rendered[12]);
        }

        [Fact]
        public void ReadTableReturnsOnlySectionDataRows()
        {
            var serializer = CreateSerializer();
            var sheetRows = new[]
            {
                new[] { "SheetBindings" },
                new[] { "SheetName", "SystemKey", "ProjectId" },
                new[] { "Sheet1", "current-business-system", "performance" },
                new[] { string.Empty, string.Empty, string.Empty },
                new[] { string.Empty, string.Empty, string.Empty },
                new[] { "SheetFieldMappings" },
                new[] { "SheetName", "HeaderId", "ApiFieldKey", "IsIdColumn" },
                new[] { "Sheet1", "row_id", "row_id", "TRUE" },
                new[] { "Sheet1", "owner_name", "owner_name", "FALSE" },
            };

            var rows = InvokeReadTable(serializer, "SheetFieldMappings", sheetRows);

            Assert.Equal(2, rows.Length);
            Assert.Equal(new[] { "Sheet1", "row_id", "row_id", "TRUE" }, rows[0]);
            Assert.Equal(new[] { "Sheet1", "owner_name", "owner_name", "FALSE" }, rows[1]);
        }

        [Fact]
        public void ReadTableReturnsEmptyWhenSectionMissing()
        {
            var serializer = CreateSerializer();
            var rows = InvokeReadTable(
                serializer,
                "SheetFieldMappings",
                new[]
                {
                    new[] { "SheetBindings" },
                    new[] { "SheetName", "SystemKey" },
                    new[] { "Sheet1", "current-business-system" },
                });

            Assert.Empty(rows);
        }

        private static object CreateSerializer()
        {
            var assembly = LoadAddInAssembly();
            var serializerType = assembly.GetType(
                "OfficeAgent.ExcelAddIn.Excel.MetadataSheetLayoutSerializer",
                throwOnError: true);
            return Activator.CreateInstance(serializerType);
        }

        private static object CreateSection(string title, string[] headers, string[][] rows)
        {
            var assembly = LoadAddInAssembly();
            var type = assembly.GetType(
                "OfficeAgent.ExcelAddIn.Excel.MetadataSectionDocument",
                throwOnError: true);
            return Activator.CreateInstance(type, title, headers, rows);
        }

        private static string[][] InvokeRender(object serializer, IDictionary<string, object> sections)
        {
            var sectionType = LoadAddInAssembly().GetType(
                "OfficeAgent.ExcelAddIn.Excel.MetadataSectionDocument",
                throwOnError: true);
            var dictionaryType = typeof(Dictionary<,>).MakeGenericType(typeof(string), sectionType);
            var typedSections = Activator.CreateInstance(dictionaryType);
            var addMethod = dictionaryType.GetMethod("Add");

            foreach (var pair in sections)
            {
                addMethod.Invoke(typedSections, new[] { pair.Key, pair.Value });
            }

            var method = serializer.GetType().GetMethod(
                "Render",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return (string[][])method.Invoke(serializer, new[] { typedSections });
        }

        private static string[][] InvokeReadTable(object serializer, string tableName, string[][] sheetRows)
        {
            var method = serializer.GetType().GetMethod(
                "ReadTable",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return (string[][])method.Invoke(serializer, new object[] { tableName, sheetRows });
        }

        private static Assembly LoadAddInAssembly()
        {
            return Assembly.LoadFrom(
                Path.GetFullPath(
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
                        "OfficeAgent.ExcelAddIn.dll")));
        }
    }
}
