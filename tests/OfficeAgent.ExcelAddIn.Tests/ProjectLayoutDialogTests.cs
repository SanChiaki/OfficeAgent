using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
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
            var args = new object[] { seed, "abc", "2", "3", CreateHostStrings(), null, null };

            var success = (bool)method.Invoke(null, args);

            Assert.False(success);
            Assert.Null(args[5]);
            Assert.Equal("HeaderStartRow 必须是正整数。", (string)args[6]);
        }

        [Fact]
        public void TryCreateBindingRejectsDataStartInsideHeaderArea()
        {
            var method = GetTryCreateBindingMethod();
            var seed = CreateSeedBinding();
            var args = new object[] { seed, "1", "2", "2", CreateHostStrings(), null, null };

            var success = (bool)method.Invoke(null, args);

            Assert.False(success);
            Assert.Null(args[5]);
            Assert.Equal(
                "DataStartRow 必须大于或等于 HeaderStartRow + HeaderRowCount。",
                (string)args[6]);
        }

        [Fact]
        public void TryCreateBindingRejectsNonNumericHeaderRowCount()
        {
            var method = GetTryCreateBindingMethod();
            var seed = CreateSeedBinding();
            var args = new object[] { seed, "1", "abc", "3", CreateHostStrings(), null, null };

            var success = (bool)method.Invoke(null, args);

            Assert.False(success);
            Assert.Null(args[5]);
            Assert.Equal("HeaderRowCount 必须是正整数。", (string)args[6]);
        }

        [Fact]
        public void TryCreateBindingRejectsNonNumericDataStartRow()
        {
            var method = GetTryCreateBindingMethod();
            var seed = CreateSeedBinding();
            var args = new object[] { seed, "1", "2", "abc", CreateHostStrings(), null, null };

            var success = (bool)method.Invoke(null, args);

            Assert.False(success);
            Assert.Null(args[5]);
            Assert.Equal("DataStartRow 必须是正整数。", (string)args[6]);
        }

        [Fact]
        public void TryCreateBindingReturnsEditedBindingForValidValues()
        {
            var method = GetTryCreateBindingMethod();
            var seed = CreateSeedBinding();
            var args = new object[] { seed, "4", "1", "5", CreateHostStrings(), null, null };

            var success = (bool)method.Invoke(null, args);

            Assert.True(success);
            var binding = Assert.IsType<SheetBinding>(args[5]);
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
            Assert.Null(args[6]);
        }

        [Fact]
        public void DialogLayoutDoesNotClipOrOverlapWhenFontScalesUp()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog())
                using (var scaledFont = CreateStressFont(dialog.Font))
                {
                    ApplyFont(dialog, scaledFont);
                    dialog.CreateControl();
                    dialog.PerformLayout();

                    AssertLayoutFits(dialog);
                }
            });
        }

        [Fact]
        public void DialogCanRenderEnglishChrome()
        {
            RunInSta(() =>
            {
                using (var dialog = CreateDialog("en"))
                {
                    dialog.CreateControl();
                    dialog.PerformLayout();

                    Assert.Equal("Configure sheet layout", dialog.Text);
                    Assert.Contains(FindButtons(dialog), button => string.Equals(button.Text, "OK", StringComparison.Ordinal));
                    Assert.Contains(FindButtons(dialog), button => string.Equals(button.Text, "Cancel", StringComparison.Ordinal));
                    Assert.Contains(
                        FindLabels(dialog),
                        label => label.Text?.IndexOf("Current binding:", StringComparison.Ordinal) >= 0);
                }
            });
        }

        private static MethodInfo GetTryCreateBindingMethod()
        {
            return GetProjectLayoutDialogType()
                .GetMethod(
                    "TryCreateBinding",
                    BindingFlags.Static | BindingFlags.NonPublic)
                ?? throw new InvalidOperationException("ProjectLayoutDialog.TryCreateBinding was not found.");
        }

        private static Form CreateDialog(string locale = "zh")
        {
            var hostStrings = CreateHostStrings(locale);

            return (Form)Activator.CreateInstance(
                GetProjectLayoutDialogType(),
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                args: new object[] { CreateSeedBinding(), hostStrings },
                culture: null);
        }

        private static object CreateHostStrings(string locale = "zh")
        {
            var hostStringsType = LoadAddInAssembly().GetType(
                "OfficeAgent.ExcelAddIn.Localization.HostLocalizedStrings",
                throwOnError: true);
            var forLocale = hostStringsType.GetMethod("ForLocale", BindingFlags.Public | BindingFlags.Static);

            return forLocale.Invoke(null, new object[] { locale });
        }

        private static Type GetProjectLayoutDialogType()
        {
            return LoadAddInAssembly()
                .GetType("OfficeAgent.ExcelAddIn.Dialogs.ProjectLayoutDialog", throwOnError: true);
        }

        private static Assembly LoadAddInAssembly()
        {
            return Assembly.LoadFrom(ResolveAddInAssemblyPath());
        }

        private static IEnumerable<Button> FindButtons(Control root)
        {
            return root.Controls.Cast<Control>()
                .SelectMany(control => FindButtons(control))
                .Concat(root is Button button ? new[] { button } : Array.Empty<Button>());
        }

        private static IEnumerable<Label> FindLabels(Control root)
        {
            return root.Controls.Cast<Control>()
                .SelectMany(control => FindLabels(control))
                .Concat(root is Label label ? new[] { label } : Array.Empty<Label>());
        }

        private static void AssertLayoutFits(Control root)
        {
            foreach (var parent in EnumerateParents(root))
            {
                var visibleChildren = parent.Controls.Cast<Control>().ToArray();

                foreach (var child in visibleChildren)
                {
                    Assert.True(
                        parent.ClientRectangle.Contains(child.Bounds),
                        $"Control '{child.Name ?? child.Text}' exceeds its parent bounds.");

                    if (child is Label label)
                    {
                        var preferred = label.GetPreferredSize(new Size(Math.Max(label.Width, 1), 0));
                        Assert.True(
                            preferred.Height <= label.Height,
                            $"Label '{label.Text}' is clipped vertically. Preferred height: {preferred.Height}, actual height: {label.Height}.");
                    }
                }

                for (var i = 0; i < visibleChildren.Length; i++)
                {
                    for (var j = i + 1; j < visibleChildren.Length; j++)
                    {
                        Assert.False(
                            visibleChildren[i].Bounds.IntersectsWith(visibleChildren[j].Bounds),
                            $"Controls '{visibleChildren[i].Name ?? visibleChildren[i].Text}' and '{visibleChildren[j].Name ?? visibleChildren[j].Text}' overlap.");
                    }
                }
            }
        }

        private static IEnumerable<Control> EnumerateParents(Control root)
        {
            yield return root;

            foreach (Control child in root.Controls)
            {
                foreach (var descendant in EnumerateParents(child))
                {
                    yield return descendant;
                }
            }
        }

        private static void ApplyFont(Control root, Font font)
        {
            root.Font = font;

            foreach (Control child in root.Controls)
            {
                ApplyFont(child, font);
            }
        }

        private static Font CreateStressFont(Font fallbackFont)
        {
            const string ChineseUiFontName = "Microsoft YaHei UI";
            var family = FontFamily.Families.FirstOrDefault(item => string.Equals(item.Name, ChineseUiFontName, StringComparison.Ordinal))
                ?? fallbackFont.FontFamily;
            var size = Math.Max(fallbackFont.Size + 4f, 14f);
            return new Font(family, size, fallbackFont.Style);
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
