using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using OfficeAgent.Core.Models;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class AiColumnMappingPreviewDialogTests
    {
        [Fact]
        public void PreviewGridShowsOnlyActionableColumnsAndExcelLetters()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog(CreatePreview(), locale: "zh"))
                {
                    dialog.CreateControl();
                    dialog.PerformLayout();

                    var grid = FindGrid(dialog);
                    Assert.NotNull(grid);
                    Assert.Equal(
                        new[] { "是否修改", "列号", "当前表头（L1/L2）", "匹配表头（L1/L2）" },
                        grid.Columns.Cast<DataGridViewColumn>().Select(column => column.HeaderText).ToArray());
                    Assert.IsType<DataGridViewCheckBoxColumn>(grid.Columns[0]);
                    Assert.Equal("B", grid.Rows[0].Cells[1].Value);
                    Assert.Equal("基础信息 / 负责人", grid.Rows[0].Cells[2].Value);
                    Assert.Equal("负责人", grid.Rows[0].Cells[3].Value);
                    Assert.True((bool)grid.Rows[0].Cells[0].Value);
                }
            });
        }

        [Fact]
        public void ApplyDialogSelectionPersistsUncheckedRowsToPreview()
        {
            RunInSta(() =>
            {
                var preview = CreatePreview();
                using (var dialog = CreateDialog(preview, locale: "zh"))
                {
                    dialog.CreateControl();
                    dialog.PerformLayout();

                    var grid = FindGrid(dialog);
                    grid.Rows[0].Cells[0].Value = false;
                    InvokeApplySelection(dialog);

                    Assert.False(preview.Items[0].ShouldApply);
                }
            });
        }

        private static Form CreateDialog(AiColumnMappingPreview preview, string locale)
        {
            return (Form)Activator.CreateInstance(
                GetDialogType(),
                BindingFlags.Instance | BindingFlags.NonPublic,
                binder: null,
                args: new[] { preview, CreateHostStrings(locale) },
                culture: null);
        }

        private static AiColumnMappingPreview CreatePreview()
        {
            return new AiColumnMappingPreview
            {
                Items = new[]
                {
                    new AiColumnMappingPreviewItem
                    {
                        ExcelColumn = 2,
                        SuggestedExcelL1 = "基础信息",
                        SuggestedExcelL2 = "负责人",
                        TargetIsdpL1 = "负责人",
                        TargetHeaderId = "owner_name",
                        TargetApiFieldKey = "owner_name",
                        Confidence = 0.91,
                        Status = AiColumnMappingPreviewStatuses.Accepted,
                    },
                },
            };
        }

        private static void InvokeApplySelection(Form dialog)
        {
            var method = GetDialogType().GetMethod(
                "ApplySelectionToPreview",
                BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.NotNull(method);

            method.Invoke(dialog, Array.Empty<object>());
        }

        private static DataGridView FindGrid(Control root)
        {
            return root.Controls.Cast<Control>()
                .SelectMany(control => EnumerateControls(control))
                .OfType<DataGridView>()
                .FirstOrDefault();
        }

        private static IEnumerable<Control> EnumerateControls(Control root)
        {
            yield return root;

            foreach (Control child in root.Controls)
            {
                foreach (var descendant in EnumerateControls(child))
                {
                    yield return descendant;
                }
            }
        }

        private static object CreateHostStrings(string locale)
        {
            var hostStringsType = LoadAddInAssembly().GetType(
                "OfficeAgent.ExcelAddIn.Localization.HostLocalizedStrings",
                throwOnError: true);
            var forLocale = hostStringsType.GetMethod("ForLocale", BindingFlags.Public | BindingFlags.Static);

            return forLocale.Invoke(null, new object[] { locale });
        }

        private static Type GetDialogType()
        {
            return LoadAddInAssembly()
                .GetType("OfficeAgent.ExcelAddIn.Dialogs.AiColumnMappingPreviewDialog", throwOnError: true);
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

        private static void RunInSta(Action action)
        {
            Exception failure = null;
            var thread = new Thread(() =>
            {
                try
                {
                    action();
                }
                catch (Exception error)
                {
                    failure = error;
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (failure != null)
            {
                throw new TargetInvocationException(failure);
            }
        }
    }
}
