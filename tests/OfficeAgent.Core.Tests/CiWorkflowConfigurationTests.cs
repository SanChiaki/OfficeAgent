using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class CiWorkflowConfigurationTests
    {
        [Fact]
        public void BuildMsiWorkflowUsesNode24CompatibleActionVersions()
        {
            var workflowText = File.ReadAllText(ResolveRepositoryPath(
                ".github",
                "workflows",
                "build-msi.yml"));

            Assert.Contains("uses: actions/checkout@v5", workflowText, StringComparison.Ordinal);
            Assert.Contains("uses: actions/setup-node@v5", workflowText, StringComparison.Ordinal);
            Assert.Contains("uses: actions/setup-dotnet@v5", workflowText, StringComparison.Ordinal);
            Assert.Contains("uses: microsoft/setup-msbuild@v3", workflowText, StringComparison.Ordinal);
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
