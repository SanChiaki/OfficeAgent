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
            Assert.Contains("xISDP.Setup-x86.msi", bundleWxsText, StringComparison.Ordinal);
            Assert.Contains("xISDP.Setup-x64.msi", bundleWxsText, StringComparison.Ordinal);
            Assert.Contains("Compressed=\"yes\"", bundleWxsText, StringComparison.Ordinal);
        }

        [Fact]
        public void SetupBuildScriptReferencesOfflineBundleOutputsAndExtensions()
        {
            var buildScriptText = ReadRepositoryFile(
                "installer",
                "OfficeAgent.Setup",
                "build.ps1");

            Assert.Contains("xISDP.Setup.exe", buildScriptText, StringComparison.Ordinal);
            Assert.Contains("xISDP.Setup-x86.msi", buildScriptText, StringComparison.Ordinal);
            Assert.Contains("xISDP.Setup-x64.msi", buildScriptText, StringComparison.Ordinal);
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
        public void InstallerAuthoringUsesXisdpDisplayNames()
        {
            var productWxsText = ReadRepositoryFile(
                "installer",
                "OfficeAgent.Setup",
                "Product.wxs");
            var bundleWxsText = ReadRepositoryFile(
                "installer",
                "OfficeAgent.SetupBundle",
                "Bundle.wxs");

            Assert.Contains("Name=\"xISDP\"", productWxsText, StringComparison.Ordinal);
            Assert.Contains("<SummaryInformation Description=\"xISDP VSTO add-in for Excel\" />", productWxsText, StringComparison.Ordinal);
            Assert.Contains("<Feature Id=\"OfficeAgentFeature\" Title=\"xISDP\" Level=\"1\">", productWxsText, StringComparison.Ordinal);
            Assert.Contains("Name=\"xISDP\"", bundleWxsText, StringComparison.Ordinal);
            Assert.Contains("DisplayName=\"xISDP (x86)\"", bundleWxsText, StringComparison.Ordinal);
            Assert.Contains("DisplayName=\"xISDP (x64)\"", bundleWxsText, StringComparison.Ordinal);
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
        public void SetupBundlePassesPrerequisiteBypassPropertiesToMsiPackages()
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
                bundleWxsText.Split(new[] { "Name=\"SKIPVSTORUNTIMECHECK\"" }, StringSplitOptions.None).Length - 1);
            Assert.Equal(
                2,
                bundleWxsText.Split(new[] { "<MsiProperty Name=\"SKIPWEBVIEW2CHECK\" Value=\"1\" />" }, StringSplitOptions.None).Length - 1);
            Assert.Equal(
                2,
                bundleWxsText.Split(new[] { "<MsiProperty Name=\"SKIPVSTORUNTIMECHECK\" Value=\"1\" />" }, StringSplitOptions.None).Length - 1);
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
        public void BuildMsiWorkflowStagesVstoRuntimeBeforeBuildingInstaller()
        {
            var workflowText = ReadRepositoryFile(".github", "workflows", "build-msi.yml");

            var stageIndex = workflowText.IndexOf("Stage VSTO runtime redistributable", StringComparison.Ordinal);
            var buildIndex = workflowText.IndexOf("Build installer", StringComparison.Ordinal);

            Assert.True(stageIndex >= 0, "Expected the MSI workflow to stage the VSTO runtime redistributable.");
            Assert.True(buildIndex > stageIndex, "Expected the VSTO runtime redistributable to be staged before the installer build.");
            Assert.Contains("vstor_redist.exe", workflowText, StringComparison.Ordinal);
            Assert.Contains(
                "download.microsoft.com/download/5/d/2/5d24f8f8-efbb-4b63-aa33-3785e3104713/vstor_redist.exe",
                workflowText,
                StringComparison.Ordinal);
            Assert.Contains("Get-FileHash", workflowText, StringComparison.Ordinal);
            Assert.Contains(
                "CFE1A40BBE4A50022DB2164ABDB0154984E2CECB761A23CDC81CB5754F6E0A18",
                workflowText,
                StringComparison.Ordinal);
            Assert.Contains("artifacts/installer/*.exe", workflowText, StringComparison.Ordinal);
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
                "SKIPVSTORUNTIMECHECK",
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
            var deploymentNotesText = ReadRepositoryFile("docs", "vsto-deployment-notes.md");
            var releaseNotesText = ReadRepositoryFile("docs", "vsto-release-notes.md");
            var customIconsGuideText = ReadRepositoryFile("docs", "ribbon-button-custom-icons-guide.md");
            var bundleReadmeText = ReadRepositoryFile(
                "installer",
                "OfficeAgent.SetupBundle",
                "README.md");

            Assert.Contains("xISDP.Setup.exe", agentsText, StringComparison.Ordinal);
            Assert.Contains("xISDP.Setup.exe", checklistText, StringComparison.Ordinal);
            Assert.Contains("artifacts/installer/xISDP.Setup-x86.msi", deploymentNotesText, StringComparison.Ordinal);
            Assert.Contains("artifacts/installer/xISDP.Setup-x64.msi", deploymentNotesText, StringComparison.Ordinal);
            Assert.Contains("artifacts/installer/xISDP.Setup-x86.msi", releaseNotesText, StringComparison.Ordinal);
            Assert.Contains("artifacts/installer/xISDP.Setup-x64.msi", releaseNotesText, StringComparison.Ordinal);
            Assert.Contains("# xISDP Setup Bundle", bundleReadmeText, StringComparison.Ordinal);
            Assert.Contains("`xISDP` tab", customIconsGuideText, StringComparison.Ordinal);
            Assert.Contains("vstor_redist.exe", checklistText, StringComparison.Ordinal);
            Assert.DoesNotContain(
                "artifacts/installer/OfficeAgent.Setup",
                deploymentNotesText,
                StringComparison.Ordinal);
            Assert.DoesNotContain(
                "artifacts/installer/OfficeAgent.Setup",
                releaseNotesText,
                StringComparison.Ordinal);
            Assert.DoesNotContain(
                "artifacts/installer/X-ISDP.Setup",
                deploymentNotesText,
                StringComparison.Ordinal);
            Assert.DoesNotContain(
                "artifacts/installer/X-ISDP.Setup",
                releaseNotesText,
                StringComparison.Ordinal);
            Assert.DoesNotContain(
                "`ISDP` tab",
                customIconsGuideText,
                StringComparison.Ordinal);
            Assert.DoesNotContain(
                "`X-ISDP` tab",
                customIconsGuideText,
                StringComparison.Ordinal);
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
