using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Sync;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class RibbonSyncControllerTests
    {
        [Fact]
        public void NewControllerDefaultsToSelectProjectDisplayWhenNoBinding()
        {
            var controller = CreateController(new FakeSystemConnector(), new FakeWorksheetMetadataStore(), () => "Sheet1");

            Assert.Equal("先选择项目", ReadActiveProjectDisplayName(controller));
        }

        [Fact]
        public void SelectProjectSavesBindingAndUpdatesActiveProjectState()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            var controller = CreateController(new FakeSystemConnector(), metadataStore, () => "SheetA");
            var option = new ProjectOption
            {
                SystemKey = "current-business-system",
                ProjectId = "project-1",
                DisplayName = "项目一",
            };

            InvokeSelectProject(controller, option);

            Assert.NotNull(metadataStore.LastSavedBinding);
            Assert.Equal("SheetA", metadataStore.LastSavedBinding.SheetName);
            Assert.Equal("current-business-system", metadataStore.LastSavedBinding.SystemKey);
            Assert.Equal("project-1", metadataStore.LastSavedBinding.ProjectId);
            Assert.Equal("项目一", metadataStore.LastSavedBinding.ProjectName);
            Assert.Equal("项目一", ReadActiveProjectDisplayName(controller));
            Assert.Equal("project-1", ReadActiveProjectId(controller));
        }

        [Fact]
        public void RefreshActiveProjectFromSheetMetadataLoadsBindingForCurrentSheet()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            metadataStore.Bindings["SheetWithBinding"] = new SheetBinding
            {
                SheetName = "SheetWithBinding",
                SystemKey = "current-business-system",
                ProjectId = "project-2",
                ProjectName = "项目二",
            };

            var controller = CreateController(new FakeSystemConnector(), metadataStore, () => "SheetWithBinding");

            InvokeRefresh(controller);

            Assert.Equal("项目二", ReadActiveProjectDisplayName(controller));
            Assert.Equal("project-2", ReadActiveProjectId(controller));
        }

        [Fact]
        public void RefreshActiveProjectFromSheetMetadataFallsBackToDefaultWhenBindingMissing()
        {
            var metadataStore = new FakeWorksheetMetadataStore();
            var controller = CreateController(new FakeSystemConnector(), metadataStore, () => "SheetWithoutBinding");

            InvokeRefresh(controller);

            Assert.Equal("先选择项目", ReadActiveProjectDisplayName(controller));
            Assert.Equal(string.Empty, ReadActiveProjectId(controller));
        }

        private static object CreateController(
            ISystemConnector connector,
            IWorksheetMetadataStore metadataStore,
            Func<string> sheetNameProvider)
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var controllerType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.RibbonSyncController", throwOnError: true);
            var syncService = new WorksheetSyncService(
                connector,
                metadataStore,
                new WorksheetChangeTracker(),
                new SyncOperationPreviewFactory());

            var ctor = controllerType.GetConstructor(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                types: new[] { typeof(ISystemConnector), typeof(IWorksheetMetadataStore), typeof(WorksheetSyncService), typeof(Func<string>) },
                modifiers: null);

            if (ctor == null)
            {
                throw new InvalidOperationException("RibbonSyncController constructor was not found.");
            }

            return ctor.Invoke(new object[] { connector, metadataStore, syncService, sheetNameProvider });
        }

        private static void InvokeSelectProject(object controller, ProjectOption option)
        {
            var method = controller.GetType().GetMethod(
                "SelectProject",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                types: new[] { typeof(ProjectOption) },
                modifiers: null);

            if (method == null)
            {
                throw new InvalidOperationException("RibbonSyncController.SelectProject(ProjectOption) was not found.");
            }

            method.Invoke(controller, new object[] { option });
        }

        private static void InvokeRefresh(object controller)
        {
            var method = controller.GetType().GetMethod(
                "RefreshActiveProjectFromSheetMetadata",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("RibbonSyncController.RefreshActiveProjectFromSheetMetadata() was not found.");
            }

            method.Invoke(controller, null);
        }

        private static string ReadActiveProjectDisplayName(object controller)
        {
            return (string)controller.GetType().GetProperty(
                "ActiveProjectDisplayName",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).GetValue(controller);
        }

        private static string ReadActiveProjectId(object controller)
        {
            return (string)controller.GetType().GetProperty(
                "ActiveProjectId",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).GetValue(controller);
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

        private sealed class FakeSystemConnector : ISystemConnector
        {
            public IReadOnlyList<ProjectOption> GetProjects()
            {
                return Array.Empty<ProjectOption>();
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

        private sealed class FakeWorksheetMetadataStore : IWorksheetMetadataStore
        {
            public Dictionary<string, SheetBinding> Bindings { get; } = new Dictionary<string, SheetBinding>(StringComparer.OrdinalIgnoreCase);

            public SheetBinding LastSavedBinding { get; private set; }

            public void SaveBinding(SheetBinding binding)
            {
                LastSavedBinding = binding;
                Bindings[binding.SheetName] = binding;
            }

            public SheetBinding LoadBinding(string sheetName)
            {
                if (!Bindings.TryGetValue(sheetName, out var binding))
                {
                    throw new InvalidOperationException("No binding.");
                }

                return binding;
            }

            public WorksheetSnapshotCell[] LoadSnapshot(string sheetName)
            {
                return Array.Empty<WorksheetSnapshotCell>();
            }

            public void SaveSnapshot(string sheetName, WorksheetSnapshotCell[] cells)
            {
                throw new NotSupportedException();
            }
        }
    }
}
