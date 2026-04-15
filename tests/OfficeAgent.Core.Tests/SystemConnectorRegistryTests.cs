using System;
using System.Collections.Generic;
using System.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class SystemConnectorRegistryTests
    {
        [Fact]
        public void GetProjectsAggregatesProjectsFromAllRegisteredConnectors()
        {
            var registry = new SystemConnectorRegistry(new ISystemConnector[]
            {
                new FakeSystemConnector(
                    "system-a",
                    new[]
                    {
                        CreateProject("system-a", "project-1", "项目一"),
                    }),
                new FakeSystemConnector(
                    "system-b",
                    new[]
                    {
                        CreateProject("system-b", "project-1", "另一个项目一"),
                        CreateProject("system-b", "project-2", "项目二"),
                    }),
            });

            var projects = registry.GetProjects();

            Assert.Equal(3, projects.Count);
            Assert.Contains(projects, project => project.SystemKey == "system-a" && project.ProjectId == "project-1");
            Assert.Contains(projects, project => project.SystemKey == "system-b" && project.ProjectId == "project-1");
            Assert.Contains(projects, project => project.SystemKey == "system-b" && project.ProjectId == "project-2");
        }

        [Fact]
        public void GetRequiredConnectorReturnsConnectorForMatchingSystemKey()
        {
            var connector = new FakeSystemConnector("system-a", Array.Empty<ProjectOption>());
            var registry = new SystemConnectorRegistry(new[] { connector });

            var resolved = registry.GetRequiredConnector("system-a");

            Assert.Same(connector, resolved);
        }

        [Fact]
        public void GetProjectsFillsMissingProjectSystemKeyFromConnector()
        {
            var registry = new SystemConnectorRegistry(new ISystemConnector[]
            {
                new FakeSystemConnector(
                    "system-a",
                    new[]
                    {
                        CreateProject(string.Empty, "project-1", "项目一"),
                    }),
            });

            var project = Assert.Single(registry.GetProjects());

            Assert.Equal("system-a", project.SystemKey);
            Assert.Equal("project-1", project.ProjectId);
        }

        private static ProjectOption CreateProject(string systemKey, string projectId, string displayName)
        {
            return new ProjectOption
            {
                SystemKey = systemKey,
                ProjectId = projectId,
                DisplayName = displayName,
            };
        }

        private sealed class FakeSystemConnector : ISystemConnector
        {
            private readonly IReadOnlyList<ProjectOption> projects;

            public FakeSystemConnector(string systemKey, IReadOnlyList<ProjectOption> projects)
            {
                SystemKey = systemKey;
                this.projects = projects ?? Array.Empty<ProjectOption>();
            }

            public string SystemKey { get; }

            public IReadOnlyList<ProjectOption> GetProjects()
            {
                return projects;
            }

            public SheetBinding CreateBindingSeed(string sheetName, ProjectOption project)
            {
                throw new NotSupportedException();
            }

            public FieldMappingTableDefinition GetFieldMappingDefinition(string projectId)
            {
                throw new NotSupportedException();
            }

            public IReadOnlyList<SheetFieldMappingRow> BuildFieldMappingSeed(string sheetName, string projectId)
            {
                throw new NotSupportedException();
            }

            public WorksheetSchema GetSchema(string projectId)
            {
                throw new NotSupportedException();
            }

            public IReadOnlyList<IDictionary<string, object>> Find(string projectId, IReadOnlyList<string> rowIds, IReadOnlyList<string> fieldKeys)
            {
                throw new NotSupportedException();
            }

            public void BatchSave(string projectId, IReadOnlyList<CellChange> changes)
            {
                throw new NotSupportedException();
            }
        }
    }
}
