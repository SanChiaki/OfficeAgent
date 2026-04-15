using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class MockServerFixtureConfigurationTests
    {
        [Fact]
        public void IntegrationTestsShareSingleMockServerCollection()
        {
            var businessApiTestText = File.ReadAllText(ResolveRepositoryPath(
                "tests",
                "OfficeAgent.IntegrationTests",
                "BusinessApiIntegrationTests.cs"));
            var connectorTestText = File.ReadAllText(ResolveRepositoryPath(
                "tests",
                "OfficeAgent.IntegrationTests",
                "CurrentBusinessSystemConnectorIntegrationTests.cs"));

            Assert.Contains("[Collection(MockServerCollection.Name)]", businessApiTestText, StringComparison.Ordinal);
            Assert.Contains("[Collection(MockServerCollection.Name)]", connectorTestText, StringComparison.Ordinal);
            Assert.Contains("[CollectionDefinition(MockServerCollection.Name)]", businessApiTestText, StringComparison.Ordinal);
        }

        [Fact]
        public void MockServerFixtureWaitsForBothSsoAndBusinessPorts()
        {
            var businessApiTestText = File.ReadAllText(ResolveRepositoryPath(
                "tests",
                "OfficeAgent.IntegrationTests",
                "BusinessApiIntegrationTests.cs"));

            Assert.Contains("WaitForServerReady(client, SsoUrl + \"/login\")", businessApiTestText, StringComparison.Ordinal);
            Assert.Contains("WaitForServerReady(client, BusinessUrl + \"/projects\")", businessApiTestText, StringComparison.Ordinal);
        }

        [Fact]
        public void MockServerFixtureDoesNotAbortStartupPollingOnProcessHasExited()
        {
            var businessApiTestText = File.ReadAllText(ResolveRepositoryPath(
                "tests",
                "OfficeAgent.IntegrationTests",
                "BusinessApiIntegrationTests.cs"));

            Assert.DoesNotContain("if (process.HasExited)", businessApiTestText, StringComparison.Ordinal);
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
