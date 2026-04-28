using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class InstallerAuthoringTests
    {
        private static readonly XNamespace WixNamespace = "http://wixtoolset.org/schemas/v4/wxs";

        [Fact]
        public void ProductWxsRunsAppSearchBeforeLaunchConditionsInBothSequences()
        {
            var document = LoadInstallerAuthoring();
            var package = document.Root?.Element(WixNamespace + "Package");

            Assert.NotNull(package);

            var installUiSequence = package?.Elements(WixNamespace + "InstallUISequence").SingleOrDefault();
            var installExecuteSequence = package?.Elements(WixNamespace + "InstallExecuteSequence").SingleOrDefault();

            Assert.NotNull(installUiSequence);
            Assert.NotNull(installExecuteSequence);
            Assert.Equal("LaunchConditions", installUiSequence?.Element(WixNamespace + "AppSearch")?.Attribute("Before")?.Value);
            Assert.Equal("LaunchConditions", installExecuteSequence?.Element(WixNamespace + "AppSearch")?.Attribute("Before")?.Value);
        }

        [Fact]
        public void ProductWxsSearchesBothRegistryViewsForVstoRuntimeAndWebView2()
        {
            var document = LoadInstallerAuthoring();
            var package = document.Root?.Element(WixNamespace + "Package");

            Assert.NotNull(package);

            var properties = package
                ?.Elements(WixNamespace + "Property")
                .ToDictionary(property => property.Attribute("Id")?.Value ?? string.Empty);

            AssertRegistrySearch(
                properties?["VSTORUNTIMEVERSION"],
                @"SOFTWARE\Microsoft\VSTO Runtime Setup\v4R",
                "Version",
                "always64");
            AssertRegistrySearch(
                properties?["VSTORUNTIMEWOW6432VERSION"],
                @"SOFTWARE\Microsoft\VSTO Runtime Setup\v4R",
                "Version",
                "always32");
            AssertRegistrySearch(
                properties?["VSTORUNTIMEINSTALL"],
                @"SOFTWARE\Microsoft\VSTO Runtime Setup\v4",
                "Install",
                "always64");
            AssertRegistrySearch(
                properties?["VSTORUNTIMEWOW6432INSTALL"],
                @"SOFTWARE\Microsoft\VSTO Runtime Setup\v4",
                "Install",
                "always32");
            AssertRegistrySearch(
                properties?["WEBVIEW2RUNTIMENATIVE"],
                @"SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
                "pv",
                "always64");
            AssertRegistrySearch(
                properties?["WEBVIEW2RUNTIMEWOW6432"],
                @"SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}",
                "pv",
                "always32");
        }

        [Fact]
        public void ProductWxsUsesXisdpProductName()
        {
            var document = LoadInstallerAuthoring();
            var package = document.Root?.Element(WixNamespace + "Package");

            Assert.NotNull(package);
            Assert.Equal("xISDP", package?.Attribute("Name")?.Value);
            Assert.Equal(
                "xISDP VSTO add-in for Excel",
                package?.Element(WixNamespace + "SummaryInformation")?.Attribute("Description")?.Value);
        }

        [Fact]
        public void ProductWxsLaunchConditionAcceptsEitherVstoRuntimeSignal()
        {
            var document = LoadInstallerAuthoring();
            var vstoLaunchCondition = document
                .Descendants(WixNamespace + "Launch")
                .First(element => (element.Attribute("Message")?.Value ?? string.Empty).Contains("VSTO runtime"));
            var condition = vstoLaunchCondition.Attribute("Condition")?.Value ?? string.Empty;

            Assert.Contains("VSTORUNTIMEVERSION", condition);
            Assert.Contains("VSTORUNTIMEWOW6432VERSION", condition);
            Assert.Contains("VSTORUNTIMEINSTALL", condition);
            Assert.Contains("VSTORUNTIMEWOW6432INSTALL", condition);
        }

        private static XDocument LoadInstallerAuthoring()
        {
            var productWxsPath = Path.GetFullPath(Path.Combine(
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
                "installer",
                "OfficeAgent.Setup",
                "Product.wxs"));

            return XDocument.Load(productWxsPath);
        }

        private static void AssertRegistrySearch(XElement property, string expectedKey, string expectedName, string expectedBitness)
        {
            Assert.NotNull(property);

            var search = property?.Element(WixNamespace + "RegistrySearch");
            Assert.NotNull(search);
            Assert.Equal(expectedKey, search?.Attribute("Key")?.Value);
            Assert.Equal(expectedName, search?.Attribute("Name")?.Value);
            Assert.Equal(expectedBitness, search?.Attribute("Bitness")?.Value);
        }
    }
}
