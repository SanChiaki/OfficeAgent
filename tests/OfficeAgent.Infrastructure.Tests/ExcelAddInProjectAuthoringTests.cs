using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeAgent.Infrastructure.Tests
{
    public sealed class ExcelAddInProjectAuthoringTests
    {
        [Fact]
        public void OfficeAgentExcelAddInKeepsManifestSigningEnabled()
        {
            var document = LoadProject("src", "OfficeAgent.ExcelAddIn", "OfficeAgent.ExcelAddIn.csproj");
            var signManifestsElements = document
                .Root?
                .Descendants()
                .Where(element => element.Name.LocalName == "SignManifests")
                .ToArray();

            Assert.NotNull(signManifestsElements);
            Assert.Contains(
                signManifestsElements,
                element => string.IsNullOrEmpty(element.Attribute("Condition")?.Value) &&
                           string.Equals(element.Value.Trim(), "true", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(
                signManifestsElements,
                element => string.Equals(element.Value.Trim(), "false", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void ExcelAddInTestsBuildTargetUsesSharedSignedVstoBuildScript()
        {
            var helperScriptPath = GetRepositoryPath("eng", "Build-VstoAddIn.ps1");
            var document = LoadProject("tests", "OfficeAgent.ExcelAddIn.Tests", "OfficeAgent.ExcelAddIn.Tests.csproj");
            var buildTarget = document
                .Root?
                .Descendants()
                .FirstOrDefault(element =>
                    element.Name.LocalName == "Target" &&
                    string.Equals(element.Attribute("Name")?.Value, "BuildOfficeAddIn", StringComparison.Ordinal));
            var execCommands = buildTarget?
                .Descendants()
                .Where(element => element.Name.LocalName == "Exec")
                .Select(element => element.Attribute("Command")?.Value ?? string.Empty)
                .ToArray();

            Assert.True(File.Exists(helperScriptPath), $"Expected shared helper script at {helperScriptPath}.");
            Assert.NotNull(execCommands);
            Assert.Contains(execCommands, command => command.IndexOf("Build-VstoAddIn.ps1", StringComparison.Ordinal) >= 0);
        }

        [Fact]
        public void InstallerBuildUsesSharedSignedVstoBuildScript()
        {
            var helperScriptPath = GetRepositoryPath("eng", "Build-VstoAddIn.ps1");
            var installerBuildScript = File.ReadAllText(GetRepositoryPath("installer", "OfficeAgent.Setup", "build.ps1"));

            Assert.True(File.Exists(helperScriptPath), $"Expected shared helper script at {helperScriptPath}.");
            Assert.True(
                installerBuildScript.IndexOf("Build-VstoAddIn.ps1", StringComparison.Ordinal) >= 0,
                "Expected installer build script to reuse the shared VSTO build helper.");
        }

        private static XDocument LoadProject(params string[] relativePathSegments)
        {
            return XDocument.Load(GetRepositoryPath(relativePathSegments));
        }

        private static string GetRepositoryPath(params string[] relativePathSegments)
        {
            return Path.GetFullPath(Path.Combine(
                AppContext.BaseDirectory,
                "..",
                "..",
                "..",
                "..",
                "..",
                Path.Combine(relativePathSegments)));
        }
    }
}
