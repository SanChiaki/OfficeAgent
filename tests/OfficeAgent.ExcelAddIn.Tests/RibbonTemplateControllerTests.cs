using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class RibbonTemplateControllerTests
    {
        [Fact]
        public void RefreshActiveTemplateStateWithAdHocSheetEnablesApplyAndSaveAsButDisablesSave()
        {
            var catalog = new FakeTemplateCatalog();
            catalog.StateBySheet["Sheet1"] = new SheetTemplateState
            {
                HasProjectBinding = true,
                CanApplyTemplate = true,
                CanSaveAsTemplate = true,
                CanSaveTemplate = false,
                TemplateOrigin = "ad-hoc",
            };
            var dialogs = new FakeTemplateDialogService();
            var controller = CreateController(catalog, dialogs, () => "Sheet1");

            InvokeRefresh(controller);

            Assert.True(ReadBoolProperty(controller, "CanApplyTemplate"));
            Assert.True(ReadBoolProperty(controller, "CanSaveAsTemplate"));
            Assert.False(ReadBoolProperty(controller, "CanSaveTemplate"));
            Assert.Equal("未绑定模板", ReadStringProperty(controller, "ActiveTemplateDisplayName"));
        }

        [Fact]
        public void ExecuteApplyTemplateUsesSelectedTemplateAndRefreshesState()
        {
            var catalog = new FakeTemplateCatalog();
            catalog.StateBySheet["Sheet1"] = new SheetTemplateState
            {
                HasProjectBinding = true,
                CanApplyTemplate = true,
                CanSaveAsTemplate = true,
                ProjectDisplayName = "绩效项目",
            };
            catalog.TemplatesBySheet["Sheet1"] = new[]
            {
                new TemplateDefinition { TemplateId = "tpl-performance-a", TemplateName = "条件A" },
            };
            var dialogs = new FakeTemplateDialogService
            {
                SelectedTemplateId = "tpl-performance-a",
            };
            var controller = CreateController(catalog, dialogs, () => "Sheet1");

            InvokeExecuteApplyTemplate(controller);

            Assert.Equal("tpl-performance-a", catalog.LastAppliedTemplateId);
            Assert.Contains(dialogs.InfoMessages, message => message.IndexOf("条件A", StringComparison.Ordinal) >= 0);
        }

        [Fact]
        public void ExecuteSaveTemplateRoutesRevisionConflictToSaveAsWhenUserChoosesNo()
        {
            var catalog = new FakeTemplateCatalog
            {
                SaveConflictMessage = "Template revision conflict.",
            };
            catalog.StateBySheet["Sheet1"] = new SheetTemplateState
            {
                HasProjectBinding = true,
                CanSaveTemplate = true,
                CanSaveAsTemplate = true,
                TemplateId = "tpl-performance-a",
                TemplateName = "条件A",
                TemplateRevision = 1,
                StoredTemplateRevision = 3,
                TemplateOrigin = "store-template",
            };
            var dialogs = new FakeTemplateDialogService
            {
                RevisionConflictResultName = "No",
                SaveAsTemplateName = "条件B",
            };
            var controller = CreateController(catalog, dialogs, () => "Sheet1");

            InvokeExecuteSaveTemplate(controller);

            Assert.Equal(("Sheet1", "tpl-performance-a", 1, false), catalog.SaveExistingCalls[0]);
            Assert.Equal("条件B", catalog.LastSavedAsTemplateName);
        }

        [Fact]
        public void ExecuteSaveTemplateOverwritesWhenUserChoosesYes()
        {
            var catalog = new FakeTemplateCatalog
            {
                SaveConflictMessage = "模板版本已变化。",
            };
            catalog.StateBySheet["Sheet1"] = new SheetTemplateState
            {
                HasProjectBinding = true,
                CanSaveTemplate = true,
                TemplateId = "tpl-performance-a",
                TemplateName = "条件A",
                TemplateRevision = 2,
                StoredTemplateRevision = 4,
                TemplateOrigin = "store-template",
            };
            var dialogs = new FakeTemplateDialogService
            {
                RevisionConflictResultName = "Yes",
            };
            var controller = CreateController(catalog, dialogs, () => "Sheet1");

            InvokeExecuteSaveTemplate(controller);

            Assert.Equal(2, catalog.SaveExistingCalls.Count);
            Assert.Equal(("Sheet1", "tpl-performance-a", 2, false), catalog.SaveExistingCalls[0]);
            Assert.Equal(("Sheet1", "tpl-performance-a", 2, true), catalog.SaveExistingCalls[1]);
            Assert.Contains(dialogs.InfoMessages, message => message.IndexOf("覆盖模板完成", StringComparison.Ordinal) >= 0);
        }

        [Fact]
        public void ExecuteSaveAsTemplateUsesSuggestedCopyName()
        {
            var catalog = new FakeTemplateCatalog();
            catalog.StateBySheet["Sheet1"] = new SheetTemplateState
            {
                HasProjectBinding = true,
                CanSaveAsTemplate = true,
                TemplateName = "条件A",
                TemplateOrigin = "store-template",
            };
            var dialogs = new FakeTemplateDialogService
            {
                SaveAsTemplateName = "条件A-副本2",
            };
            var controller = CreateController(catalog, dialogs, () => "Sheet1");

            InvokeExecuteSaveAsTemplate(controller);

            Assert.Equal("条件A-副本", dialogs.LastSuggestedTemplateName);
            Assert.Equal("条件A-副本2", catalog.LastSavedAsTemplateName);
            Assert.Contains(dialogs.InfoMessages, message => message.IndexOf("另存模板完成", StringComparison.Ordinal) >= 0);
        }

        private static object CreateController(
            FakeTemplateCatalog catalog,
            FakeTemplateDialogService dialogService,
            Func<string> sheetNameProvider)
        {
            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var controllerType = addInAssembly.GetType("OfficeAgent.ExcelAddIn.RibbonTemplateController", throwOnError: true);
            var dialogInterface = addInAssembly.GetType("OfficeAgent.ExcelAddIn.Dialogs.IRibbonTemplateDialogService", throwOnError: true);

            var ctor = controllerType.GetConstructor(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                types: new[]
                {
                    typeof(ITemplateCatalog),
                    typeof(Func<string>),
                    dialogInterface,
                },
                modifiers: null);

            if (ctor == null)
            {
                throw new InvalidOperationException("RibbonTemplateController constructor was not found.");
            }

            return ctor.Invoke(new object[] { catalog, sheetNameProvider, dialogService.GetTransparentProxy() });
        }

        private static void InvokeRefresh(object controller)
        {
            var method = controller.GetType().GetMethod(
                "RefreshActiveTemplateStateFromSheetMetadata",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("RibbonTemplateController.RefreshActiveTemplateStateFromSheetMetadata() was not found.");
            }

            method.Invoke(controller, null);
        }

        private static void InvokeExecuteApplyTemplate(object controller)
        {
            var method = controller.GetType().GetMethod(
                "ExecuteApplyTemplate",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("RibbonTemplateController.ExecuteApplyTemplate() was not found.");
            }

            method.Invoke(controller, null);
        }

        private static void InvokeExecuteSaveTemplate(object controller)
        {
            var method = controller.GetType().GetMethod(
                "ExecuteSaveTemplate",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("RibbonTemplateController.ExecuteSaveTemplate() was not found.");
            }

            method.Invoke(controller, null);
        }

        private static void InvokeExecuteSaveAsTemplate(object controller)
        {
            var method = controller.GetType().GetMethod(
                "ExecuteSaveAsTemplate",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

            if (method == null)
            {
                throw new InvalidOperationException("RibbonTemplateController.ExecuteSaveAsTemplate() was not found.");
            }

            method.Invoke(controller, null);
        }

        private static bool ReadBoolProperty(object controller, string propertyName)
        {
            return (bool)controller.GetType().GetProperty(
                propertyName,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).GetValue(controller);
        }

        private static string ReadStringProperty(object controller, string propertyName)
        {
            return (string)controller.GetType().GetProperty(
                propertyName,
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

        private sealed class FakeTemplateCatalog : ITemplateCatalog
        {
            public Dictionary<string, SheetTemplateState> StateBySheet { get; } =
                new Dictionary<string, SheetTemplateState>(StringComparer.OrdinalIgnoreCase);

            public Dictionary<string, IReadOnlyList<TemplateDefinition>> TemplatesBySheet { get; } =
                new Dictionary<string, IReadOnlyList<TemplateDefinition>>(StringComparer.OrdinalIgnoreCase);

            public List<(string SheetName, string TemplateId, int ExpectedRevision, bool Overwrite)> SaveExistingCalls { get; } =
                new List<(string SheetName, string TemplateId, int ExpectedRevision, bool Overwrite)>();

            public string LastAppliedTemplateId { get; private set; } = string.Empty;

            public string LastSavedAsTemplateName { get; private set; } = string.Empty;

            public string SaveConflictMessage { get; set; } = string.Empty;

            public IReadOnlyList<TemplateDefinition> ListTemplates(string sheetName)
            {
                return TemplatesBySheet.TryGetValue(sheetName, out var templates)
                    ? templates
                    : Array.Empty<TemplateDefinition>();
            }

            public SheetTemplateState GetSheetState(string sheetName)
            {
                return StateBySheet.TryGetValue(sheetName, out var state)
                    ? state
                    : new SheetTemplateState();
            }

            public void ApplyTemplateToSheet(string sheetName, string templateId)
            {
                LastAppliedTemplateId = templateId;
            }

            public void SaveSheetToExistingTemplate(string sheetName, string templateId, int expectedRevision, bool overwriteRevisionConflict)
            {
                SaveExistingCalls.Add((sheetName, templateId, expectedRevision, overwriteRevisionConflict));
                if (!string.IsNullOrWhiteSpace(SaveConflictMessage) && !overwriteRevisionConflict)
                {
                    throw new InvalidOperationException(SaveConflictMessage);
                }
            }

            public void SaveSheetAsNewTemplate(string sheetName, string templateName)
            {
                LastSavedAsTemplateName = templateName;
            }

            public void DetachTemplate(string sheetName)
            {
            }
        }

        private sealed class FakeTemplateDialogService : RealProxy
        {
            public FakeTemplateDialogService()
                : base(LoadDialogInterfaceType())
            {
            }

            public List<string> InfoMessages { get; } = new List<string>();

            public List<string> WarningMessages { get; } = new List<string>();

            public List<string> ErrorMessages { get; } = new List<string>();

            public string SelectedTemplateId { get; set; } = string.Empty;

            public string SaveAsTemplateName { get; set; } = string.Empty;

            public string LastSuggestedTemplateName { get; private set; } = string.Empty;

            public bool ConfirmApplyTemplateOverwriteResult { get; set; } = true;

            public string RevisionConflictResultName { get; set; } = "Cancel";

            public override IMessage Invoke(IMessage msg)
            {
                var call = (IMethodCallMessage)msg;

                switch (call.MethodName)
                {
                    case "ShowTemplatePicker":
                        return new ReturnMessage(SelectedTemplateId, null, 0, call.LogicalCallContext, call);
                    case "ShowSaveAsTemplateDialog":
                        LastSuggestedTemplateName = (string)call.InArgs[0];
                        return new ReturnMessage(SaveAsTemplateName, null, 0, call.LogicalCallContext, call);
                    case "ConfirmApplyTemplateOverwrite":
                        return new ReturnMessage(ConfirmApplyTemplateOverwriteResult, null, 0, call.LogicalCallContext, call);
                    case "ShowTemplateRevisionConflictDialog":
                        var returnType = ((MethodInfo)call.MethodBase).ReturnType;
                        return new ReturnMessage(
                            Enum.Parse(returnType, RevisionConflictResultName, ignoreCase: false),
                            null,
                            0,
                            call.LogicalCallContext,
                            call);
                    case "ShowInfo":
                        InfoMessages.Add((string)call.InArgs[0]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "ShowWarning":
                        WarningMessages.Add((string)call.InArgs[0]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    case "ShowError":
                        ErrorMessages.Add((string)call.InArgs[0]);
                        return new ReturnMessage(null, null, 0, call.LogicalCallContext, call);
                    default:
                        throw new NotSupportedException(call.MethodName);
                }
            }

            public new object GetTransparentProxy()
            {
                return base.GetTransparentProxy();
            }

            private static Type LoadDialogInterfaceType()
            {
                return Assembly.LoadFrom(ResolveAddInAssemblyPath())
                    .GetType("OfficeAgent.ExcelAddIn.Dialogs.IRibbonTemplateDialogService", throwOnError: true);
            }
        }
    }
}
