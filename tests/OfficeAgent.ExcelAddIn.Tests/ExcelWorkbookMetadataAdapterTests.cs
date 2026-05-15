using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class ExcelWorkbookMetadataAdapterTests
    {
        [Fact]
        public void EnsureWorksheetRestoresTheOriginalActiveWorksheetWhenMetadataSheetIsCreated()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());
            var ensureWorksheet = adapterType.GetMethod("EnsureWorksheet", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            ensureWorksheet.Invoke(adapter, new object[] { "xISDP_Setting", true });

            Assert.Equal("BusinessSheet", application.ActiveSheet.Name);
            Assert.NotNull(application.MetadataSheet);
            Assert.Equal("xISDP_Setting", application.MetadataSheet.Name);
            Assert.Equal(1, application.BusinessSheet.ActivateCount);
        }

        [Fact]
        public void ReadTableRenamesLegacyMetadataSheetToCanonicalName()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            application.CreateWorksheet("ISDP_Setting");
            application.MetadataSheet.SetCell(1, 1, "SheetBindings");
            application.MetadataSheet.SetCell(2, 1, "SheetName");
            application.MetadataSheet.SetCell(2, 2, "SystemKey");
            application.MetadataSheet.SetCell(3, 1, "Sheet1");
            application.MetadataSheet.SetCell(3, 2, "current-business-system");

            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

            var rows = (string[][])adapterType.GetMethod("ReadTable").Invoke(adapter, new object[] { "SheetBindings" });

            Assert.Single(rows);
            Assert.Equal(new[] { "Sheet1", "current-business-system" }, rows[0]);
            Assert.Equal("xISDP_Setting", application.MetadataSheet.Name);
            Assert.Equal("BusinessSheet", application.ActiveSheet.Name);
        }

        [Fact]
        public void ReadHeadersReturnsSectionHeadersAndRestoresActiveWorksheet()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            application.CreateWorksheet("xISDP_Setting");
            application.MetadataSheet.SetCell(1, 1, "SheetFieldMappings");
            application.MetadataSheet.SetCell(2, 1, "SheetName");
            application.MetadataSheet.SetCell(2, 2, "HeaderType");
            application.MetadataSheet.SetCell(2, 3, "ISDP L1");
            application.MetadataSheet.SetCell(2, 4, "Excel L1");

            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

            var headers = (string[])adapterType.GetMethod("ReadHeaders").Invoke(adapter, new object[] { "SheetFieldMappings" });

            Assert.Equal(new[] { "SheetName", "HeaderType", "ISDP L1", "Excel L1" }, headers);
            Assert.Equal("BusinessSheet", application.ActiveSheet.Name);
        }

        [Fact]
        public void WriteTableRenamesLegacyMetadataSheetBeforeRewriting()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            application.CreateWorksheet("ISDP_Setting");
            application.MetadataSheet.SetCell(1, 1, "TemplateBindings");
            application.MetadataSheet.SetCell(2, 1, "SheetName");
            application.MetadataSheet.SetCell(3, 1, "Sheet1");

            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

            adapterType.GetMethod("WriteTable").Invoke(adapter, new object[]
            {
                "SheetBindings",
                new[] { "SheetName", "SystemKey" },
                new[] { new[] { "Sheet1", "current-business-system" } },
            });

            Assert.Equal("xISDP_Setting", application.MetadataSheet.Name);
            Assert.Equal("SheetBindings", application.MetadataSheet.GetCell(1, 1));
            Assert.Equal("TemplateBindings", application.MetadataSheet.GetCell(6, 1));
        }

        [Fact]
        public void ReadTablePrefersCanonicalMetadataSheetWhenLegacyAlsoExists()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            var canonicalSheet = application.CreateWorksheet("xISDP_Setting");
            canonicalSheet.SetCell(1, 1, "SheetBindings");
            canonicalSheet.SetCell(2, 1, "SheetName");
            canonicalSheet.SetCell(2, 2, "SystemKey");
            canonicalSheet.SetCell(3, 1, "CanonicalSheet");
            canonicalSheet.SetCell(3, 2, "canonical-system");
            var legacySheet = application.CreateWorksheet("ISDP_Setting");
            legacySheet.SetCell(1, 1, "SheetBindings");
            legacySheet.SetCell(2, 1, "SheetName");
            legacySheet.SetCell(2, 2, "SystemKey");
            legacySheet.SetCell(3, 1, "LegacySheet");
            legacySheet.SetCell(3, 2, "legacy-system");

            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

            var rows = (string[][])adapterType.GetMethod("ReadTable").Invoke(adapter, new object[] { "SheetBindings" });

            var row = Assert.Single(rows);
            Assert.Equal("CanonicalSheet", row[0]);
            Assert.Equal("xISDP_Setting", canonicalSheet.Name);
            Assert.Equal("ISDP_Setting", legacySheet.Name);
        }

        [Fact]
        public void WriteTableRewritesMetadataSheetUsingTitledSections()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

            adapterType.GetMethod("WriteTable").Invoke(adapter, new object[]
            {
                "SheetBindings",
                new[] { "SheetName", "SystemKey" },
                new[] { new[] { "Sheet1", "current-business-system" } },
            });

            adapterType.GetMethod("WriteTable").Invoke(adapter, new object[]
            {
                "SheetFieldMappings",
                new[] { "SheetName", "HeaderId", "ApiFieldKey" },
                new[] { new[] { "Sheet1", "row_id", "row_id" } },
            });

            Assert.Equal("SheetBindings", application.MetadataSheet.GetCell(1, 1));
            Assert.Equal("SheetName", application.MetadataSheet.GetCell(2, 1));
            Assert.Equal("Sheet1", application.MetadataSheet.GetCell(3, 1));
            Assert.Equal(string.Empty, application.MetadataSheet.GetCell(4, 1));
            Assert.Equal(string.Empty, application.MetadataSheet.GetCell(5, 1));
            Assert.Equal("SheetFieldMappings", application.MetadataSheet.GetCell(6, 1));
            Assert.Equal("HeaderId", application.MetadataSheet.GetCell(7, 2));
            Assert.Equal("row_id", application.MetadataSheet.GetCell(8, 2));
        }

        [Fact]
        public void WriteTablePreservesExistingTemplateBindingsWhenRewritingAnotherSection()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            application.CreateWorksheet("xISDP_Setting");
            application.MetadataSheet.SetCell(1, 1, "TemplateBindings");
            application.MetadataSheet.SetCell(2, 1, "SheetName");
            application.MetadataSheet.SetCell(2, 2, "TemplateId");
            application.MetadataSheet.SetCell(2, 3, "TemplateOrigin");
            application.MetadataSheet.SetCell(3, 1, "Sheet1");
            application.MetadataSheet.SetCell(3, 2, "tpl-performance-a");
            application.MetadataSheet.SetCell(3, 3, "store-template");
            application.MetadataSheet.SetCell(6, 1, "SheetBindings");
            application.MetadataSheet.SetCell(7, 1, "SheetName");
            application.MetadataSheet.SetCell(7, 2, "SystemKey");
            application.MetadataSheet.SetCell(8, 1, "Sheet1");
            application.MetadataSheet.SetCell(8, 2, "current-business-system");

            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

            adapterType.GetMethod("WriteTable").Invoke(adapter, new object[]
            {
                "SheetBindings",
                new[] { "SheetName", "SystemKey" },
                new[] { new[] { "Sheet1", "current-business-system" } },
            });

            Assert.Equal("SheetBindings", application.MetadataSheet.GetCell(1, 1));
            Assert.Equal("current-business-system", application.MetadataSheet.GetCell(3, 2));
            Assert.Equal("TemplateBindings", application.MetadataSheet.GetCell(6, 1));
            Assert.Equal("tpl-performance-a", application.MetadataSheet.GetCell(8, 2));
        }

        [Fact]
        public void WriteTableCanRewriteMetadataSheetUsingRangeValue2WithoutDirectCellWrites()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

            application.CreateWorksheet("xISDP_Setting");
            application.MetadataSheet.ThrowOnDirectCellWrites = true;

            adapterType.GetMethod("WriteTable").Invoke(adapter, new object[]
            {
                "SheetFieldMappings",
                new[] { "SheetName", "HeaderId", "ApiFieldKey" },
                new[]
                {
                    new[] { "Sheet1", "row_id", "row_id" },
                    new[] { "Sheet1", "owner_name", "owner_name" },
                },
            });

            Assert.Equal(1, application.MetadataSheet.RangeValue2WriteCount);
            Assert.Equal("SheetFieldMappings", application.MetadataSheet.GetCell(1, 1));
            Assert.Equal("HeaderId", application.MetadataSheet.GetCell(2, 2));
            Assert.Equal("owner_name", application.MetadataSheet.GetCell(4, 2));
        }

        [Fact]
        public void ReadTableReadsRowsBackFromTitledSections()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            application.CreateWorksheet("xISDP_Setting");
            application.MetadataSheet.SetCell(1, 1, "SheetBindings");
            application.MetadataSheet.SetCell(2, 1, "SheetName");
            application.MetadataSheet.SetCell(2, 2, "SystemKey");
            application.MetadataSheet.SetCell(3, 1, "Sheet1");
            application.MetadataSheet.SetCell(3, 2, "current-business-system");

            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

            var rows = (string[][])adapterType.GetMethod("ReadTable").Invoke(adapter, new object[] { "SheetBindings" });

            Assert.Single(rows);
            Assert.Equal(new[] { "Sheet1", "current-business-system" }, rows[0]);
        }

        [Fact]
        public void ReadTableCanParseSectionsFromUsedRangeValue2WithoutDirectCellReads()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            application.CreateWorksheet("xISDP_Setting");
            application.MetadataSheet.SetCell(1, 1, "SheetBindings");
            application.MetadataSheet.SetCell(2, 1, "SheetName");
            application.MetadataSheet.SetCell(2, 2, "SystemKey");
            application.MetadataSheet.SetCell(3, 1, "Sheet1");
            application.MetadataSheet.SetCell(3, 2, "current-business-system");
            application.MetadataSheet.ThrowOnDirectCellReads = true;

            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

            var rows = (string[][])adapterType.GetMethod("ReadTable").Invoke(adapter, new object[] { "SheetBindings" });

            Assert.Single(rows);
            Assert.Equal(new[] { "Sheet1", "current-business-system" }, rows[0]);
            Assert.Equal(1, application.MetadataSheet.UsedRangeValue2ReadCount);
        }

        [Fact]
        public void ReadTableDoesNotCreateSettingsWorksheetWhenMetadataSheetIsMissing()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

            var rows = (string[][])adapterType.GetMethod("ReadTable").Invoke(adapter, new object[] { "SheetBindings" });

            Assert.Empty(rows);
            Assert.Null(application.MetadataSheet);
            Assert.Equal("BusinessSheet", application.ActiveSheet.Name);
        }

        [Fact]
        public void ApplyMetadataPresentationFormatsEachSheetNameHeaderRowWithinFirstFiftyRows()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            application.CreateWorksheet("xISDP_Setting");
            application.MetadataSheet.SetCell(1, 1, "TemplateBindings");
            application.MetadataSheet.SetCell(2, 1, "SheetName");
            application.MetadataSheet.SetCell(3, 1, "SalesSheet");
            application.MetadataSheet.SetCell(4, 1, "sheetname");
            application.MetadataSheet.SetCell(5, 1, "ProjectSheet");

            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

            adapterType.GetMethod("ApplyMetadataPresentation").Invoke(adapter, new object[] { "xISDP_Setting", false });

            Assert.True(application.MetadataSheet.GetRowBold(2));
            Assert.True(application.MetadataSheet.GetRowBold(4));
            Assert.False(application.MetadataSheet.GetRowBold(3));
            Assert.False(application.MetadataSheet.GetRowBold(5));
        }

        [Fact]
        public void ApplyMetadataPresentationClearsPreviousBlueBoldRowsBeforeFormattingCurrentSheetNameHeaders()
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var excelAssembly = LoadExcelInteropAssembly();
            var worksheetType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Worksheet", throwOnError: true);
            var sheetsType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Sheets", throwOnError: true);
            var workbookType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Workbook", throwOnError: true);
            var applicationType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Application", throwOnError: true);
            var rangeType = excelAssembly.GetType("Microsoft.Office.Interop.Excel.Range", throwOnError: true);

            var application = new LayoutAwareFakeExcelApplication(applicationType, workbookType, sheetsType, worksheetType, rangeType);
            application.CreateWorksheet("xISDP_Setting");
            application.MetadataSheet.SetCell(1, 1, "SheetBindings");
            application.MetadataSheet.SetCell(2, 1, "OtherHeader");
            application.MetadataSheet.SetCell(3, 1, "SheetName");
            application.MetadataSheet.GetOrCreateRowState(2).Bold = true;
            application.MetadataSheet.GetOrCreateRowState(2).Color = ColorTranslator.ToOle(Color.Blue);

            var adapterType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Excel.ExcelWorkbookMetadataAdapter", throwOnError: true);
            var adapter = Activator.CreateInstance(adapterType, application.GetTransparentProxy());

            adapterType.GetMethod("ApplyMetadataPresentation").Invoke(adapter, new object[] { "xISDP_Setting", false });

            Assert.False(application.MetadataSheet.GetRowBold(2));
            Assert.Null(application.MetadataSheet.GetRowColor(2));
            Assert.Contains("Automatic", Convert.ToString(application.MetadataSheet.GetRowColorIndex(2)) ?? string.Empty, StringComparison.OrdinalIgnoreCase);
            Assert.True(application.MetadataSheet.GetRowBold(3));
        }

        private static Assembly LoadExcelInteropAssembly()
        {
            try
            {
                return Assembly.Load("Microsoft.Office.Interop.Excel");
            }
            catch (FileNotFoundException)
            {
                var candidatePaths = new[]
                {
                    Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                        "Microsoft Visual Studio",
                        "Shared",
                        "Visual Studio Tools for Office",
                        "PIA",
                        "Office15",
                        "Microsoft.Office.Interop.Excel.dll"),
                    @"E:\VSgongxiang\Visual Studio Tools for Office\PIA\Office15\Microsoft.Office.Interop.Excel.dll",
                };

                var fallbackPath = candidatePaths.FirstOrDefault(File.Exists);
                if (string.IsNullOrWhiteSpace(fallbackPath))
                {
                    throw;
                }

                return Assembly.LoadFrom(fallbackPath);
            }
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

        private sealed class LayoutAwareFakeExcelApplication : RealProxy
        {
            private readonly LayoutAwareFakeExcelWorkbook workbook;
            private readonly Type applicationType;

            public LayoutAwareFakeExcelApplication(Type applicationType, Type workbookType, Type sheetsType, Type worksheetType, Type rangeType)
                : base(applicationType)
            {
                this.applicationType = applicationType;
                BusinessSheet = new LayoutAwareFakeExcelWorksheet(worksheetType, rangeType, this, "BusinessSheet");
                workbook = new LayoutAwareFakeExcelWorkbook(workbookType, sheetsType, worksheetType, rangeType, this, BusinessSheet);
                ActiveSheet = BusinessSheet;
            }

            public LayoutAwareFakeExcelWorksheet BusinessSheet { get; }

            public LayoutAwareFakeExcelWorksheet MetadataSheet { get; set; }

            public LayoutAwareFakeExcelWorksheet ActiveSheet { get; set; }

            public LayoutAwareFakeExcelWorksheet CreateWorksheet(string name)
            {
                var worksheet = workbook.Worksheets.AddWorksheet(name);
                if (string.Equals(name, "xISDP_Setting", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, "ISDP_Setting", StringComparison.OrdinalIgnoreCase))
                {
                    MetadataSheet = worksheet;
                }

                return worksheet;
            }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                return call.MethodName switch
                {
                    "GetType" => new ReturnMessage(applicationType, null, 0, call.LogicalCallContext, call),
                    "get_ActiveWorkbook" => new ReturnMessage(workbook.GetTransparentProxy(), null, 0, call.LogicalCallContext, call),
                    "get_ActiveSheet" => new ReturnMessage(ActiveSheet?.GetTransparentProxy(), null, 0, call.LogicalCallContext, call),
                    _ => throw new NotSupportedException(call.MethodName),
                };
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }
        }

        private sealed class LayoutAwareFakeExcelWorkbook : RealProxy
        {
            private readonly Type workbookType;

            public LayoutAwareFakeExcelWorkbook(
                Type workbookType,
                Type sheetsType,
                Type worksheetType,
                Type rangeType,
                LayoutAwareFakeExcelApplication application,
                LayoutAwareFakeExcelWorksheet initialSheet)
                : base(workbookType)
            {
                this.workbookType = workbookType;
                Worksheets = new LayoutAwareFakeExcelSheets(sheetsType, worksheetType, rangeType, application, initialSheet);
            }

            public LayoutAwareFakeExcelSheets Worksheets { get; }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                return call.MethodName switch
                {
                    "GetType" => new ReturnMessage(workbookType, null, 0, call.LogicalCallContext, call),
                    "get_Worksheets" => new ReturnMessage(Worksheets.GetTransparentProxy(), null, 0, call.LogicalCallContext, call),
                    _ => throw new NotSupportedException(call.MethodName),
                };
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }
        }

        private sealed class LayoutAwareFakeExcelSheets : RealProxy
        {
            private readonly List<LayoutAwareFakeExcelWorksheet> sheets = new List<LayoutAwareFakeExcelWorksheet>();
            private readonly Type sheetsType;
            private readonly Type worksheetType;
            private readonly Type rangeType;
            private readonly LayoutAwareFakeExcelApplication application;

            public LayoutAwareFakeExcelSheets(
                Type sheetsType,
                Type worksheetType,
                Type rangeType,
                LayoutAwareFakeExcelApplication application,
                LayoutAwareFakeExcelWorksheet initialSheet)
                : base(sheetsType)
            {
                this.sheetsType = sheetsType;
                this.worksheetType = worksheetType;
                this.rangeType = rangeType;
                this.application = application;
                sheets.Add(initialSheet);
            }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                return call.MethodName switch
                {
                    "GetType" => new ReturnMessage(sheetsType, null, 0, call.LogicalCallContext, call),
                    "get_Count" => new ReturnMessage(sheets.Count, null, 0, call.LogicalCallContext, call),
                    "get_Item" => new ReturnMessage(ResolveSheet(call.InArgs[0]).GetTransparentProxy(), null, 0, call.LogicalCallContext, call),
                    "get__Default" => new ReturnMessage(ResolveSheet(call.InArgs[0]).GetTransparentProxy(), null, 0, call.LogicalCallContext, call),
                    "Item" => new ReturnMessage(ResolveSheet(call.InArgs[0]).GetTransparentProxy(), null, 0, call.LogicalCallContext, call),
                    "Add" => HandleAdd(call),
                    "GetEnumerator" => new ReturnMessage(GetEnumerator(), null, 0, call.LogicalCallContext, call),
                    _ => throw new NotSupportedException(call.MethodName),
                };
            }

            public LayoutAwareFakeExcelWorksheet AddWorksheet(string name)
            {
                var worksheet = new LayoutAwareFakeExcelWorksheet(worksheetType, rangeType, application, name);
                sheets.Add(worksheet);
                return worksheet;
            }

            private IMessage HandleAdd(IMethodCallMessage call)
            {
                var added = AddWorksheet("Sheet" + (sheets.Count + 1));
                application.MetadataSheet = added;
                application.ActiveSheet = added;
                return new ReturnMessage(added.GetTransparentProxy(), null, 0, call.LogicalCallContext, call);
            }

            private IEnumerator GetEnumerator()
            {
                foreach (var sheet in sheets)
                {
                    yield return sheet.GetTransparentProxy();
                }
            }

            private LayoutAwareFakeExcelWorksheet ResolveSheet(object index)
            {
                var ordinal = Convert.ToInt32(index);
                return sheets[ordinal - 1];
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }
        }

        private sealed class LayoutAwareFakeExcelWorksheet : RealProxy
        {
            private readonly Dictionary<(int Row, int Column), string> cells =
                new Dictionary<(int Row, int Column), string>();
            private readonly LayoutAwareFakeExcelApplication application;
            private readonly Type worksheetType;
            private readonly Type rangeType;
            private readonly Type fontType;

            public LayoutAwareFakeExcelWorksheet(Type worksheetType, Type rangeType, LayoutAwareFakeExcelApplication application, string name)
                : base(worksheetType)
            {
                this.worksheetType = worksheetType;
                this.rangeType = rangeType;
                fontType = rangeType.Assembly.GetType("Microsoft.Office.Interop.Excel.Font", throwOnError: false) ?? typeof(object);
                this.application = application;
                Name = name;
            }

            public string Name { get; private set; }

            public int ActivateCount { get; private set; }

            public bool ThrowOnDirectCellReads { get; set; }

            public bool ThrowOnDirectCellWrites { get; set; }

            public int UsedRangeValue2ReadCount { get; private set; }

            public int RangeValue2WriteCount { get; private set; }

            public Dictionary<int, LayoutAwareFakeExcelRowState> Rows { get; } =
                new Dictionary<int, LayoutAwareFakeExcelRowState>();

            public const int AutomaticColorIndex = -4105;

            public Type FontType => fontType;

            public void SetCell(int row, int column, string value)
            {
                if (string.IsNullOrEmpty(value))
                {
                    cells.Remove((row, column));
                    return;
                }

                cells[(row, column)] = value;
            }

            public string GetCell(int row, int column)
            {
                if (ThrowOnDirectCellReads)
                {
                    throw new InvalidOperationException("Direct cell reads are disabled for this test.");
                }

                return cells.TryGetValue((row, column), out var value) ? value : string.Empty;
            }

            public object GetUsedRangeValue2()
            {
                var usedRange = GetUsedRange();
                if (!usedRange.HasValue)
                {
                    return null;
                }

                UsedRangeValue2ReadCount++;
                var firstRow = usedRange.Value.FirstRow;
                var rowCount = usedRange.Value.RowCount;
                var columnCount = usedRange.Value.ColumnCount;

                if (rowCount == 1 && columnCount == 1)
                {
                    return cells.TryGetValue((firstRow, 1), out var scalar) ? scalar : null;
                }

                var values = new object[rowCount, columnCount];
                for (var rowOffset = 0; rowOffset < rowCount; rowOffset++)
                {
                    for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
                    {
                        cells.TryGetValue((firstRow + rowOffset, columnIndex + 1), out var value);
                        values[rowOffset, columnIndex] = value;
                    }
                }

                return values;
            }

            public void ClearAllCells()
            {
                cells.Clear();
            }

            public void SetRangeValues(int startRow, int startColumn, object value, int rowCount, int columnCount)
            {
                RangeValue2WriteCount++;
                if (value is Array matrix && matrix.Rank == 2)
                {
                    var rowLowerBound = matrix.GetLowerBound(0);
                    var columnLowerBound = matrix.GetLowerBound(1);
                    for (var rowOffset = 0; rowOffset < rowCount; rowOffset++)
                    {
                        for (var columnOffset = 0; columnOffset < columnCount; columnOffset++)
                        {
                            var cellValue = matrix.GetValue(rowOffset + rowLowerBound, columnOffset + columnLowerBound);
                            SetCell(
                                startRow + rowOffset,
                                startColumn + columnOffset,
                                Convert.ToString(cellValue) ?? string.Empty);
                        }
                    }

                    return;
                }

                var scalarValue = Convert.ToString(value) ?? string.Empty;
                for (var rowOffset = 0; rowOffset < rowCount; rowOffset++)
                {
                    for (var columnOffset = 0; columnOffset < columnCount; columnOffset++)
                    {
                        SetCell(startRow + rowOffset, startColumn + columnOffset, scalarValue);
                    }
                }
            }

            public (int FirstRow, int RowCount, int ColumnCount)? GetUsedRange()
            {
                if (cells.Count == 0)
                {
                    return null;
                }

                var firstRow = cells.Keys.Min(key => key.Row);
                var lastRow = cells.Keys.Max(key => key.Row);
                var lastColumn = cells.Keys.Max(key => key.Column);
                return (firstRow, lastRow - firstRow + 1, lastColumn);
            }

            public LayoutAwareFakeExcelRowState GetOrCreateRowState(int rowIndex)
            {
                if (!Rows.TryGetValue(rowIndex, out var state))
                {
                    state = new LayoutAwareFakeExcelRowState();
                    Rows[rowIndex] = state;
                }

                return state;
            }

            public bool GetRowBold(int rowIndex)
            {
                return Rows.TryGetValue(rowIndex, out var state) && state.Bold;
            }

            public object GetRowColorIndex(int rowIndex)
            {
                return Rows.TryGetValue(rowIndex, out var state) ? state.ColorIndex : null;
            }

            public object GetRowColor(int rowIndex)
            {
                return Rows.TryGetValue(rowIndex, out var state) ? state.Color : null;
            }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                return call.MethodName switch
                {
                    "GetType" => new ReturnMessage(worksheetType, null, 0, call.LogicalCallContext, call),
                    "get_Name" => new ReturnMessage(Name, null, 0, call.LogicalCallContext, call),
                    "set_Name" => SetName(call),
                    "set_Visible" => new ReturnMessage(null, null, 0, call.LogicalCallContext, call),
                    "Activate" => Activate(call),
                    "get_Cells" => new ReturnMessage(
                        new LayoutAwareFakeExcelRange(rangeType, this, LayoutAwareFakeExcelRangeKind.CellCollection).GetTransparentProxy(),
                        null,
                        0,
                        call.LogicalCallContext,
                        call),
                    "get_Rows" => new ReturnMessage(
                        new LayoutAwareFakeExcelRange(rangeType, this, LayoutAwareFakeExcelRangeKind.RowCollection).GetTransparentProxy(),
                        null,
                        0,
                        call.LogicalCallContext,
                        call),
                    "get_UsedRange" => GetUsedRangeMessage(call),
                    _ => throw new NotSupportedException(call.MethodName),
                };
            }

            private IMessage SetName(IMethodCallMessage call)
            {
                Name = Convert.ToString(call.InArgs[0]) ?? string.Empty;
                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
            }

            private IMessage Activate(IMethodCallMessage call)
            {
                ActivateCount++;
                application.ActiveSheet = this;
                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
            }

            private IMessage GetUsedRangeMessage(IMethodCallMessage call)
            {
                var usedRange = GetUsedRange();
                if (!usedRange.HasValue)
                {
                    return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                }

                var range = new LayoutAwareFakeExcelRange(
                    rangeType,
                    this,
                    LayoutAwareFakeExcelRangeKind.UsedRange,
                    row: usedRange.Value.FirstRow,
                    count: usedRange.Value.RowCount,
                    otherCount: usedRange.Value.ColumnCount);
                return new ReturnMessage(range.GetTransparentProxy(), null, 0, call.LogicalCallContext, call);
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }
        }

        private enum LayoutAwareFakeExcelRangeKind
        {
            CellCollection,
            Cell,
            Block,
            UsedRange,
            RowCollection,
            Row,
            Rows,
            Columns,
        }

        private sealed class LayoutAwareFakeExcelRange : RealProxy
        {
            private readonly Type rangeType;
            private readonly LayoutAwareFakeExcelWorksheet worksheet;
            private readonly LayoutAwareFakeExcelRangeKind kind;
            private readonly int row;
            private readonly int column;
            private readonly int count;
            private readonly int otherCount;

            public LayoutAwareFakeExcelRange(
                Type rangeType,
                LayoutAwareFakeExcelWorksheet worksheet,
                LayoutAwareFakeExcelRangeKind kind,
                int row = 0,
                int column = 0,
                int count = 0,
                int otherCount = 0)
                : base(rangeType)
            {
                this.rangeType = rangeType;
                this.worksheet = worksheet;
                this.kind = kind;
                this.row = row;
                this.column = column;
                this.count = count;
                this.otherCount = otherCount;
            }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                return call.MethodName switch
                {
                    "GetType" => new ReturnMessage(rangeType, null, 0, call.LogicalCallContext, call),
                    "get_Item" => HandleGetItem(call),
                    "get__Default" => HandleGetItem(call),
                    "Item" => HandleGetItem(call),
                    "get_Resize" => HandleResize(call),
                    "Resize" => HandleResize(call),
                    "get_Value2" => HandleGetValue2(call),
                    "set_Value2" => HandleSetValue(call),
                    "ClearContents" => HandleClearContents(call),
                    "get_Font" => HandleGetFont(call),
                    "set_Hidden" => HandleSetHidden(call),
                    "get_Hidden" => HandleGetHidden(call),
                    "get_Row" => new ReturnMessage(row, null, 0, call.LogicalCallContext, call),
                    "get_Rows" => new ReturnMessage(
                        new LayoutAwareFakeExcelRange(rangeType, worksheet, LayoutAwareFakeExcelRangeKind.Rows, count: count).GetTransparentProxy(),
                        null,
                        0,
                        call.LogicalCallContext,
                        call),
                    "get_Columns" => new ReturnMessage(
                        new LayoutAwareFakeExcelRange(rangeType, worksheet, LayoutAwareFakeExcelRangeKind.Columns, count: otherCount).GetTransparentProxy(),
                        null,
                        0,
                        call.LogicalCallContext,
                        call),
                    "get_Count" => new ReturnMessage(count, null, 0, call.LogicalCallContext, call),
                    _ => throw new NotSupportedException(call.MethodName),
                };
            }

            private IMessage HandleGetItem(IMethodCallMessage call)
            {
                if (kind == LayoutAwareFakeExcelRangeKind.RowCollection)
                {
                    var rowIndex = Convert.ToInt32(call.InArgs[0]);
                    var rowRange = new LayoutAwareFakeExcelRange(
                        rangeType,
                        worksheet,
                        LayoutAwareFakeExcelRangeKind.Row,
                        row: rowIndex);
                    return new ReturnMessage(rowRange.GetTransparentProxy(), null, 0, call.LogicalCallContext, call);
                }

                if (kind != LayoutAwareFakeExcelRangeKind.CellCollection)
                {
                    throw new NotSupportedException(call.MethodName + ":" + kind);
                }

                var cellRow = Convert.ToInt32(call.InArgs[0]);
                var cellColumn = Convert.ToInt32(call.InArgs[1]);
                var cell = new LayoutAwareFakeExcelRange(
                    rangeType,
                    worksheet,
                    LayoutAwareFakeExcelRangeKind.Cell,
                    row: cellRow,
                    column: cellColumn);
                return new ReturnMessage(cell.GetTransparentProxy(), null, 0, call.LogicalCallContext, call);
            }

            private IMessage HandleResize(IMethodCallMessage call)
            {
                if (kind != LayoutAwareFakeExcelRangeKind.Cell && kind != LayoutAwareFakeExcelRangeKind.Block)
                {
                    throw new NotSupportedException(call.MethodName + ":" + kind);
                }

                var rowCount = call.InArgCount >= 1 ? Convert.ToInt32(call.InArgs[0]) : 1;
                var columnCount = call.InArgCount >= 2 ? Convert.ToInt32(call.InArgs[1]) : 1;
                var resized = new LayoutAwareFakeExcelRange(
                    rangeType,
                    worksheet,
                    LayoutAwareFakeExcelRangeKind.Block,
                    row: row,
                    column: column,
                    count: rowCount,
                    otherCount: columnCount);
                return new ReturnMessage(resized.GetTransparentProxy(), null, 0, call.LogicalCallContext, call);
            }

            private IMessage HandleGetValue2(IMethodCallMessage call)
            {
                object value;
                switch (kind)
                {
                    case LayoutAwareFakeExcelRangeKind.Cell:
                        value = worksheet.GetCell(row, column);
                        break;
                    case LayoutAwareFakeExcelRangeKind.UsedRange:
                        value = worksheet.GetUsedRangeValue2();
                        break;
                    case LayoutAwareFakeExcelRangeKind.Block:
                        throw new NotSupportedException(call.MethodName + ":" + kind);
                    default:
                        throw new NotSupportedException(call.MethodName + ":" + kind);
                }

                return new ReturnMessage(value, null, 0, call.LogicalCallContext, call);
            }

            private IMessage HandleGetFont(IMethodCallMessage call)
            {
                if (kind != LayoutAwareFakeExcelRangeKind.Row)
                {
                    throw new NotSupportedException(call.MethodName + ":" + kind);
                }

                return new ReturnMessage(
                    new LayoutAwareFakeExcelFontProxy(worksheet.FontType, worksheet.GetOrCreateRowState(row)).GetTransparentProxy(),
                    null,
                    0,
                    call.LogicalCallContext,
                    call);
            }

            private IMessage HandleSetHidden(IMethodCallMessage call)
            {
                if (kind != LayoutAwareFakeExcelRangeKind.Row)
                {
                    throw new NotSupportedException(call.MethodName + ":" + kind);
                }

                worksheet.GetOrCreateRowState(row).Hidden = Convert.ToBoolean(call.InArgs[0]);
                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
            }

            private IMessage HandleGetHidden(IMethodCallMessage call)
            {
                if (kind != LayoutAwareFakeExcelRangeKind.Row)
                {
                    throw new NotSupportedException(call.MethodName + ":" + kind);
                }

                return new ReturnMessage(worksheet.GetOrCreateRowState(row).Hidden, null, 0, call.LogicalCallContext, call);
            }

            private IMessage HandleSetValue(IMethodCallMessage call)
            {
                if (kind == LayoutAwareFakeExcelRangeKind.Block)
                {
                    worksheet.SetRangeValues(row, column, call.InArgs[0], count, otherCount);
                    return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                }

                if (worksheet.ThrowOnDirectCellWrites)
                {
                    throw new InvalidOperationException("Direct cell writes are disabled for this test.");
                }

                worksheet.SetCell(row, column, Convert.ToString(call.InArgs[0]) ?? string.Empty);
                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
            }

            private IMessage HandleClearContents(IMethodCallMessage call)
            {
                if (kind == LayoutAwareFakeExcelRangeKind.CellCollection)
                {
                    worksheet.ClearAllCells();
                }
                else if (kind == LayoutAwareFakeExcelRangeKind.Cell)
                {
                    worksheet.SetCell(row, column, string.Empty);
                }
                else
                {
                    throw new NotSupportedException(call.MethodName + ":" + kind);
                }

                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }
        }

        private sealed class LayoutAwareFakeExcelRowState
        {
            public bool Bold { get; set; }

            public object Color { get; set; }

            public object ColorIndex { get; set; }

            public bool Hidden { get; set; }
        }

        private sealed class LayoutAwareFakeExcelFontProxy : RealProxy
        {
            private readonly Type fontType;
            private readonly LayoutAwareFakeExcelRowState state;

            public LayoutAwareFakeExcelFontProxy(Type fontType, LayoutAwareFakeExcelRowState state)
                : base(fontType)
            {
                this.fontType = fontType;
                this.state = state;
            }

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                return call.MethodName switch
                {
                    "GetType" => new ReturnMessage(fontType, null, 0, call.LogicalCallContext, call),
                    "set_Bold" => SetBold(call),
                    "get_Bold" => new ReturnMessage(state.Bold, null, 0, call.LogicalCallContext, call),
                    "set_Color" => SetColor(call),
                    "get_Color" => new ReturnMessage(state.Color, null, 0, call.LogicalCallContext, call),
                    "set_ColorIndex" => SetColorIndex(call),
                    "get_ColorIndex" => new ReturnMessage(state.ColorIndex, null, 0, call.LogicalCallContext, call),
                    _ => throw new NotSupportedException(call.MethodName),
                };
            }

            private IMessage SetBold(IMethodCallMessage call)
            {
                state.Bold = Convert.ToBoolean(call.InArgs[0]);
                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
            }

            private IMessage SetColor(IMethodCallMessage call)
            {
                state.Color = call.InArgs[0];
                state.ColorIndex = null;
                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
            }

            private IMessage SetColorIndex(IMethodCallMessage call)
            {
                state.ColorIndex = call.InArgs[0];
                if (IsAutomaticColorIndex(call.InArgs[0]))
                {
                    state.Color = null;
                }

                return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
            }

            private static bool IsAutomaticColorIndex(object value)
            {
                try
                {
                    return Convert.ToInt32(value) == LayoutAwareFakeExcelWorksheet.AutomaticColorIndex;
                }
                catch
                {
                    return string.Equals(
                        Convert.ToString(value),
                        "xlColorIndexAutomatic",
                        StringComparison.OrdinalIgnoreCase);
                }
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }
        }
    }
}
