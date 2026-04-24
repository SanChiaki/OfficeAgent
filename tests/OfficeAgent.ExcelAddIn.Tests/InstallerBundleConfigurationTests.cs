using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class InstallerBundleConfigurationTests
    {
        [Fact]
        public void SetupBundleWixIncludesOfflinePayloadAndCompressedCab()
        {
            var bundleWxsText = ReadRepositoryFile(
                "installer",
                "OfficeAgent.SetupBundle",
                "Bundle.wxs");

            Assert.Contains("vstor_redist.exe", bundleWxsText, StringComparison.Ordinal);
            Assert.DoesNotContain("MicrosoftEdgeWebView2RuntimeInstallerX86.exe", bundleWxsText, StringComparison.Ordinal);
            Assert.DoesNotContain("MicrosoftEdgeWebView2RuntimeInstallerX64.exe", bundleWxsText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.Setup-x86.msi", bundleWxsText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.Setup-x64.msi", bundleWxsText, StringComparison.Ordinal);
            Assert.Contains("Compressed=\"yes\"", bundleWxsText, StringComparison.Ordinal);
        }

        [Fact]
        public void SetupBuildScriptReferencesOfflineBundleOutputsAndExtensions()
        {
            var buildScriptText = ReadRepositoryFile(
                "installer",
                "OfficeAgent.Setup",
                "build.ps1");

            Assert.Contains("OfficeAgent.Setup.exe", buildScriptText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.SetupBundle", buildScriptText, StringComparison.Ordinal);
            Assert.Contains("WixToolset.Bal.wixext", buildScriptText, StringComparison.Ordinal);
            Assert.Contains("WixToolset.Util.wixext", buildScriptText, StringComparison.Ordinal);
            Assert.Contains("-ext", buildScriptText, StringComparison.Ordinal);
            Assert.DoesNotContain("MicrosoftEdgeWebView2RuntimeInstallerX86.exe", buildScriptText, StringComparison.Ordinal);
            Assert.DoesNotContain("MicrosoftEdgeWebView2RuntimeInstallerX64.exe", buildScriptText, StringComparison.Ordinal);
        }

        [Fact]
        public void SetupBundleSearchesExpectedPrerequisiteRegistryLocations()
        {
            var bundleWxsText = ReadRepositoryFile(
                "installer",
                "OfficeAgent.SetupBundle",
                "Bundle.wxs");

            Assert.Contains("SearchVstoRuntimeVersion64", bundleWxsText, StringComparison.Ordinal);
            Assert.Contains("SearchVstoRuntimeVersion32", bundleWxsText, StringComparison.Ordinal);
            Assert.Contains("SearchVstoRuntimeInstall64", bundleWxsText, StringComparison.Ordinal);
            Assert.Contains("SearchVstoRuntimeInstall32", bundleWxsText, StringComparison.Ordinal);
            Assert.DoesNotContain("SearchWebView2RuntimeMachine32", bundleWxsText, StringComparison.Ordinal);
            Assert.DoesNotContain("SearchWebView2RuntimeMachine64", bundleWxsText, StringComparison.Ordinal);
            Assert.DoesNotContain("SearchWebView2RuntimeUser", bundleWxsText, StringComparison.Ordinal);
            Assert.DoesNotContain(
                "SOFTWARE\\Microsoft\\EdgeUpdate\\Clients\\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
                bundleWxsText,
                StringComparison.Ordinal);
            Assert.DoesNotContain(
                "{F1B5D7A5-8D1A-4F84-8F6A-8F92B9A6F9D0}",
                bundleWxsText,
                StringComparison.Ordinal);
        }

        [Fact]
        public void SetupBundleDoesNotCachePrerequisiteExePackages()
        {
            var bundleWxsText = ReadRepositoryFile(
                "installer",
                "OfficeAgent.SetupBundle",
                "Bundle.wxs");

            Assert.Contains(
                "Id=\"VstoRuntime\"",
                bundleWxsText,
                StringComparison.Ordinal);
            Assert.DoesNotContain(
                "Id=\"WebView2RuntimeX86\"",
                bundleWxsText,
                StringComparison.Ordinal);
            Assert.DoesNotContain(
                "Id=\"WebView2RuntimeX64\"",
                bundleWxsText,
                StringComparison.Ordinal);
            Assert.Equal(
                1,
                bundleWxsText.Split(new[] { "Cache=\"remove\"" }, StringSplitOptions.None).Length - 1);
        }

        [Fact]
        public void SetupBundlePassesWebViewBypassPropertyToMsiPackages()
        {
            var bundleWxsText = ReadRepositoryFile(
                "installer",
                "OfficeAgent.SetupBundle",
                "Bundle.wxs");

            Assert.Equal(
                2,
                bundleWxsText.Split(new[] { "Name=\"SKIPWEBVIEW2CHECK\"" }, StringSplitOptions.None).Length - 1);
            Assert.Equal(
                2,
                bundleWxsText.Split(new[] { "Value=\"1\"" }, StringSplitOptions.None).Length - 1);
        }

        [Fact]
        public void BundlePrerequisiteReadmeIncludesOnlyVstoInstallerFilename()
        {
            var prereqReadmeText = ReadRepositoryFile(
                "installer",
                "OfficeAgent.SetupBundle",
                "prereqs",
                "README.md");

            Assert.Contains("vstor_redist.exe", prereqReadmeText, StringComparison.Ordinal);
            Assert.DoesNotContain("MicrosoftEdgeWebView2RuntimeInstallerX86.exe", prereqReadmeText, StringComparison.Ordinal);
            Assert.DoesNotContain("MicrosoftEdgeWebView2RuntimeInstallerX64.exe", prereqReadmeText, StringComparison.Ordinal);
        }

        [Fact]
        public void DirectMsiProductRetainsPrerequisiteBlockMessagesAndAllowsBundleBypass()
        {
            var productWxsText = ReadRepositoryFile(
                "installer",
                "OfficeAgent.Setup",
                "Product.wxs");

            Assert.Contains(
                "requires the Microsoft Visual Studio Tools for Office Runtime 4.0 or later",
                productWxsText,
                StringComparison.Ordinal);
            Assert.Contains(
                "Install the VSTO runtime, then run this installer again.",
                productWxsText,
                StringComparison.Ordinal);
            Assert.Contains(
                "requires the Microsoft Edge WebView2 Runtime",
                productWxsText,
                StringComparison.Ordinal);
            Assert.Contains(
                "SKIPWEBVIEW2CHECK",
                productWxsText,
                StringComparison.Ordinal);
            Assert.Contains(
                "Install the Evergreen Runtime or your offline enterprise package, then run this installer again.",
                productWxsText,
                StringComparison.Ordinal);
        }

        [Fact]
        public void AgentGuidanceDocumentsReferenceSetupExeAndUpdatedChecklistFlow()
        {
            var agentsText = ReadRepositoryFile("AGENTS.md");
            var checklistText = ReadRepositoryFile("docs", "vsto-manual-test-checklist.md");

            Assert.Contains("OfficeAgent.Setup.exe", agentsText, StringComparison.Ordinal);
            Assert.Contains("OfficeAgent.Setup.exe", checklistText, StringComparison.Ordinal);
            Assert.Contains("vstor_redist.exe", checklistText, StringComparison.Ordinal);
            Assert.DoesNotContain(
                "installs WebView2 Runtime",
                checklistText,
                StringComparison.Ordinal);
            Assert.DoesNotContain(
                "MicrosoftEdgeWebView2RuntimeInstallerX64.exe",
                checklistText,
                StringComparison.Ordinal);
        }

        private static string ReadRepositoryFile(params string[] segments)
        {
            var path = ResolveRepositoryPath(segments);
            Assert.True(File.Exists(path), $"Expected repository file to exist: {path}");
            return File.ReadAllText(path);
        }

        private static string ResolveRepositoryPath(params string[] segments)
        {
            return Path.GetFullPath(Path.Combine(new[]
            {
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
            }.Concat(segments).ToArray()));
        }
    }
}
