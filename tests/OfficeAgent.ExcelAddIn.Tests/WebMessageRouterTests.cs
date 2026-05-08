using System;
using System.IO;
using System.Reflection;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Orchestration;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Skills;
using OfficeAgent.Infrastructure.Http;
using OfficeAgent.Infrastructure.Security;
using OfficeAgent.Infrastructure.Storage;
using Xunit;

namespace OfficeAgent.ExcelAddIn.Tests
{
    public sealed class WebMessageRouterTests : IDisposable
    {
        private readonly string tempDirectory;

        public WebMessageRouterTests()
        {
            tempDirectory = Path.Combine(Path.GetTempPath(), "OfficeAgent.Router.Tests", Guid.NewGuid().ToString("N"));
        }

        [Fact]
        public void SaveSettingsRejectsMissingPayloadWithoutOverwritingStoredValues()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            settingsStore.Save(new AppSettings
            {
                ApiKey = "secret-token",
                BaseUrl = "https://api.internal.example",
                BusinessBaseUrl = "https://business.internal.example",
                Model = "gpt-5-mini",
                ApiFormat = "anthropic-messages",
            });

            var router = CreateRouter(sessionStore, settingsStore);
            var responseJson = InvokeRoute(router, "{\"type\":\"bridge.saveSettings\",\"requestId\":\"req-1\"}");
            var settingsAfter = settingsStore.Load();

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
            Assert.Equal("secret-token", settingsAfter.ApiKey);
            Assert.Equal("https://api.internal.example", settingsAfter.BaseUrl);
            Assert.Equal("https://business.internal.example", settingsAfter.BusinessBaseUrl);
            Assert.Equal("gpt-5-mini", settingsAfter.Model);
            Assert.Equal("anthropic-messages", settingsAfter.ApiFormat);
        }

        [Fact]
        public void SaveSettingsRejectsEmptyObjectPayloadWithoutOverwritingStoredValues()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            settingsStore.Save(new AppSettings
            {
                ApiKey = "secret-token",
                BaseUrl = "https://api.internal.example",
                BusinessBaseUrl = "https://business.internal.example",
                Model = "gpt-5-mini",
            });

            var router = CreateRouter(sessionStore, settingsStore);
            var responseJson = InvokeRoute(router, "{\"type\":\"bridge.saveSettings\",\"requestId\":\"req-1\",\"payload\":{}}");
            var settingsAfter = settingsStore.Load();

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
            Assert.Equal("secret-token", settingsAfter.ApiKey);
            Assert.Equal("https://api.internal.example", settingsAfter.BaseUrl);
            Assert.Equal("https://business.internal.example", settingsAfter.BusinessBaseUrl);
            Assert.Equal("gpt-5-mini", settingsAfter.Model);
        }

        [Fact]
        public void SaveSettingsRejectsPayloadWithoutApiKeyWithoutOverwritingStoredValues()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            settingsStore.Save(new AppSettings
            {
                ApiKey = "secret-token",
                BaseUrl = "https://api.internal.example",
                BusinessBaseUrl = "https://business.internal.example",
                Model = "gpt-5-mini",
            });

            var router = CreateRouter(sessionStore, settingsStore);
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.saveSettings\",\"requestId\":\"req-1\",\"payload\":{\"baseUrl\":\"https://api.internal.example\",\"businessBaseUrl\":\"https://business.internal.example\",\"model\":\"gpt-5-mini\"}}");
            var settingsAfter = settingsStore.Load();

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
            Assert.Equal("secret-token", settingsAfter.ApiKey);
            Assert.Equal("https://api.internal.example", settingsAfter.BaseUrl);
            Assert.Equal("https://business.internal.example", settingsAfter.BusinessBaseUrl);
            Assert.Equal("gpt-5-mini", settingsAfter.Model);
        }

        [Fact]
        public void SaveSettingsRoundTripsBusinessBaseUrl()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());

            var router = CreateRouter(sessionStore, settingsStore);
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.saveSettings\",\"requestId\":\"req-1\",\"payload\":{\"apiKey\":\"secret-token\",\"baseUrl\":\"https://llm.internal.example\",\"businessBaseUrl\":\"https://business.internal.example\",\"model\":\"gpt-5-mini\",\"apiFormat\":\"anthropic-messages\",\"ssoUrl\":\"\",\"ssoLoginSuccessPath\":\"\"}}");

            Assert.Contains("\"ok\":true", responseJson);
            Assert.Contains("\"businessBaseUrl\":\"https://business.internal.example\"", responseJson);
            Assert.Contains("\"apiFormat\":\"anthropic-messages\"", responseJson);

            var settingsAfter = settingsStore.Load();
            Assert.Equal("https://llm.internal.example", settingsAfter.BaseUrl);
            Assert.Equal("https://business.internal.example", settingsAfter.BusinessBaseUrl);
            Assert.Equal("anthropic-messages", settingsAfter.ApiFormat);
        }

        [Fact]
        public void SaveSettingsRoundTripsUiLanguageOverride()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());

            var router = CreateRouter(sessionStore, settingsStore, resolvedUiLocale: "en");
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.saveSettings\",\"requestId\":\"req-1\",\"payload\":{\"apiKey\":\"secret-token\",\"baseUrl\":\"https://llm.internal.example\",\"businessBaseUrl\":\"https://business.internal.example\",\"model\":\"gpt-5-mini\",\"uiLanguageOverride\":\"zh\",\"ssoUrl\":\"\",\"ssoLoginSuccessPath\":\"\"}}");

            Assert.Contains("\"ok\":true", responseJson);
            Assert.Contains("\"uiLanguageOverride\":\"zh\"", responseJson);

            var settingsAfter = settingsStore.Load();
            Assert.Equal("zh", settingsAfter.UiLanguageOverride);
        }

        [Fact]
        public void SaveSettingsAcceptsNullUiLanguageOverrideAndNormalizesToSystem()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());

            var router = CreateRouter(sessionStore, settingsStore, resolvedUiLocale: "en");
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.saveSettings\",\"requestId\":\"req-1\",\"payload\":{\"apiKey\":\"secret-token\",\"baseUrl\":\"https://llm.internal.example\",\"businessBaseUrl\":\"https://business.internal.example\",\"model\":\"gpt-5-mini\",\"uiLanguageOverride\":null,\"ssoUrl\":\"\",\"ssoLoginSuccessPath\":\"\"}}");

            Assert.Contains("\"ok\":true", responseJson);
            Assert.Contains("\"uiLanguageOverride\":\"system\"", responseJson);

            var settingsAfter = settingsStore.Load();
            Assert.Equal("system", settingsAfter.UiLanguageOverride);
        }

        [Fact]
        public void GetHostContextReturnsResolvedLocaleAndPersistedOverride()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            settingsStore.Save(new AppSettings
            {
                ApiKey = "secret-token",
                BaseUrl = "https://api.internal.example",
                BusinessBaseUrl = "https://business.internal.example",
                Model = "gpt-5-mini",
                UiLanguageOverride = "zh",
            });

            var router = CreateRouter(sessionStore, settingsStore, resolvedUiLocale: "en");
            var responseJson = InvokeRoute(router, "{\"type\":\"bridge.getHostContext\",\"requestId\":\"req-1\"}");

            Assert.Contains("\"ok\":true", responseJson);
            Assert.Contains("\"resolvedUiLocale\":\"en\"", responseJson);
            Assert.Contains("\"uiLanguageOverride\":\"zh\"", responseJson);
        }

        [Fact]
        public void GetHostContextResolvesLocaleFromTheSameSettingsSnapshot()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            settingsStore.Save(new AppSettings
            {
                ApiKey = "secret-token",
                BaseUrl = "https://api.internal.example",
                BusinessBaseUrl = "https://business.internal.example",
                Model = "gpt-5-mini",
                UiLanguageOverride = "zh",
            });

            AppSettings capturedSettings = null;
            var router = CreateRouter(sessionStore, settingsStore, settings =>
            {
                capturedSettings = settings;
                return "en";
            });

            var responseJson = InvokeRoute(router, "{\"type\":\"bridge.getHostContext\",\"requestId\":\"req-1\"}");

            Assert.Contains("\"ok\":true", responseJson);
            Assert.NotNull(capturedSettings);
            Assert.Equal("zh", capturedSettings.UiLanguageOverride);
        }

        [Fact]
        public void GetHostContextRejectsPayload()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());

            var router = CreateRouter(sessionStore, settingsStore, resolvedUiLocale: "en");
            var responseJson = InvokeRoute(router, "{\"type\":\"bridge.getHostContext\",\"requestId\":\"req-1\",\"payload\":{}}");

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
        }

        [Fact]
        public void GetSettingsRejectsEmptyObjectPayload()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());

            var router = CreateRouter(sessionStore, settingsStore);
            var responseJson = InvokeRoute(router, "{\"type\":\"bridge.getSettings\",\"requestId\":\"req-1\",\"payload\":{}}");

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
        }

        [Fact]
        public void GetSelectionContextReturnsCurrentSelectionContext()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            var selectionContextService = new FakeExcelContextService(new SelectionContext
            {
                WorkbookName = "Quarterly Report.xlsx",
                SheetName = "Sheet1",
                Address = "A1:C4",
                RowCount = 4,
                ColumnCount = 3,
                IsContiguous = true,
                HeaderPreview = new[] { "Name", "Region", "Amount" },
                SampleRows = new[]
                {
                    new[] { "Project A", "CN", "42" },
                    new[] { "Project B", "US", "36" },
                },
                WarningMessage = null,
            });

            var router = CreateRouter(
                sessionStore,
                settingsStore,
                selectionContextService,
                new FakeExcelCommandExecutor());
            var responseJson = InvokeRoute(router, "{\"type\":\"bridge.getSelectionContext\",\"requestId\":\"req-1\"}");

            Assert.Contains("\"ok\":true", responseJson);
            Assert.Contains("\"workbookName\":\"Quarterly Report.xlsx\"", responseJson);
            Assert.Contains("\"sheetName\":\"Sheet1\"", responseJson);
            Assert.Contains("\"address\":\"A1:C4\"", responseJson);
        }

        [Fact]
        public void ExecuteExcelCommandExecutesReadCommandsImmediately()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            var executor = new FakeExcelCommandExecutor
            {
                ExecuteResult = new ExcelCommandResult
                {
                    CommandType = ExcelCommandTypes.ReadSelectionTable,
                    RequiresConfirmation = false,
                    Status = "completed",
                    Message = "Read selection from Sheet1 A1:C4.",
                },
            };

            var router = CreateRouter(sessionStore, settingsStore, new FakeExcelContextService(SelectionContext.Empty("No selection available.")), executor);
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.executeExcelCommand\",\"requestId\":\"req-1\",\"payload\":{\"commandType\":\"excel.readSelectionTable\",\"confirmed\":false}}");

            Assert.Contains("\"ok\":true", responseJson);
            Assert.Contains("\"requiresConfirmation\":false", responseJson);
            Assert.Equal(1, executor.ExecuteCalls);
            Assert.Equal(0, executor.PreviewCalls);
        }

        [Fact]
        public void ExecuteExcelCommandReturnsPreviewForUnconfirmedWriteCommands()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            var executor = new FakeExcelCommandExecutor
            {
                PreviewResult = new ExcelCommandResult
                {
                    CommandType = ExcelCommandTypes.AddWorksheet,
                    RequiresConfirmation = true,
                    Status = "preview",
                    Message = "Confirm worksheet creation before Excel is modified.",
                    Preview = new ExcelCommandPreview
                    {
                        Title = "Confirm Excel action",
                        Summary = "Add worksheet \"Summary\"",
                    },
                },
            };

            var router = CreateRouter(sessionStore, settingsStore, new FakeExcelContextService(SelectionContext.Empty("No selection available.")), executor);
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.executeExcelCommand\",\"requestId\":\"req-1\",\"payload\":{\"commandType\":\"excel.addWorksheet\",\"newSheetName\":\"Summary\",\"confirmed\":false}}");

            Assert.Contains("\"ok\":true", responseJson);
            Assert.Contains("\"requiresConfirmation\":true", responseJson);
            Assert.Contains("\"summary\":\"Add worksheet", responseJson);
            Assert.Equal(0, executor.ExecuteCalls);
            Assert.Equal(1, executor.PreviewCalls);
        }

        [Fact]
        public void ExecuteExcelCommandRejectsConflictingWriteRangeSheetNames()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            var executor = new FakeExcelCommandExecutor();

            var router = CreateRouter(sessionStore, settingsStore, new FakeExcelContextService(SelectionContext.Empty("No selection available.")), executor);
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.executeExcelCommand\",\"requestId\":\"req-1\",\"payload\":{\"commandType\":\"excel.writeRange\",\"sheetName\":\"Sheet1\",\"targetAddress\":\"Sheet2!A1:B2\",\"values\":[[\"Name\",\"Region\"]],\"confirmed\":false}}");

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"invalid_command\"", responseJson);
            Assert.Equal(0, executor.ExecuteCalls);
            Assert.Equal(0, executor.PreviewCalls);
        }

        [Fact]
        public void RunSkillRoutesUploadRequestsThroughTheAgentOrchestrator()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            var orchestrator = new FakeAgentOrchestrator
            {
                Result = new AgentCommandResult
                {
                    Route = AgentRouteTypes.Skill,
                    SkillName = SkillNames.UploadData,
                    RequiresConfirmation = true,
                    Status = "preview",
                    Message = "Review the upload payload before sending it to 项目A.",
                    UploadPreview = new UploadPreview
                    {
                        ProjectName = "项目A",
                    },
                },
            };

            var router = CreateRouter(
                sessionStore,
                settingsStore,
                new FakeExcelContextService(SelectionContext.Empty("No selection available.")),
                new FakeExcelCommandExecutor(),
                orchestrator);
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.runSkill\",\"requestId\":\"req-1\",\"payload\":{\"userInput\":\"把选中数据上传到项目A\",\"confirmed\":false}}");

            Assert.Contains("\"ok\":true", responseJson);
            Assert.Contains("\"route\":\"skill\"", responseJson);
            Assert.Contains("\"skillName\":\"upload_data\"", responseJson);
            Assert.Equal("把选中数据上传到项目A", orchestrator.LastEnvelope.UserInput);
        }

        [Fact]
        public void RunAgentRoutesNaturalLanguageRequestsThroughThePlannerDispatch()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            var orchestrator = new FakeAgentOrchestrator
            {
                Result = new AgentCommandResult
                {
                    Route = AgentRouteTypes.Plan,
                    RequiresConfirmation = true,
                    Status = "preview",
                    Message = "I prepared a plan. Review it before Excel is changed.",
                    Planner = new PlannerResponse
                    {
                        Mode = PlannerResponseModes.Plan,
                        AssistantMessage = "I prepared a plan. Review it before Excel is changed.",
                        Plan = new AgentPlan
                        {
                            Summary = "Create a Summary sheet.",
                            Steps = new[]
                            {
                                new AgentPlanStep
                                {
                                    Type = ExcelCommandTypes.AddWorksheet,
                                },
                            },
                        },
                    },
                },
            };

            var router = CreateRouter(
                sessionStore,
                settingsStore,
                new FakeExcelContextService(SelectionContext.Empty("No selection available.")),
                new FakeExcelCommandExecutor(),
                orchestrator);
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.runAgent\",\"requestId\":\"req-1\",\"payload\":{\"userInput\":\"Create a summary sheet\",\"confirmed\":false}}");

            Assert.Contains("\"ok\":true", responseJson);
            Assert.Contains("\"route\":\"plan\"", responseJson);
            Assert.Contains("\"mode\":\"plan\"", responseJson);
            Assert.Equal(AgentDispatchModes.Agent, orchestrator.LastEnvelope.DispatchMode);
            Assert.Equal("Create a summary sheet", orchestrator.LastEnvelope.UserInput);
        }

        [Fact]
        public void RunAgentRejectsMissingPayload()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());

            var router = CreateRouter(
                sessionStore,
                settingsStore,
                new FakeExcelContextService(SelectionContext.Empty("No selection available.")),
                new FakeExcelCommandExecutor(),
                new FakeAgentOrchestrator());
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.runAgent\",\"requestId\":\"req-1\"}");

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"malformed_payload\"", responseJson);
        }

        [Fact]
        public void ExecuteExcelCommandReturnsInternalErrorForUnexpectedExecutorFailures()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            var executor = new FakeExcelCommandExecutor
            {
                ExecuteException = new Exception("boom"),
            };

            var router = CreateRouter(
                sessionStore,
                settingsStore,
                new FakeExcelContextService(SelectionContext.Empty("No selection available.")),
                executor);
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.executeExcelCommand\",\"requestId\":\"req-1\",\"payload\":{\"commandType\":\"excel.readSelectionTable\",\"confirmed\":false}}");

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"internal_error\"", responseJson);
        }

        [Fact]
        public void RunSkillReturnsInternalErrorForUnexpectedOrchestratorFailures()
        {
            var sessionStore = new FileSessionStore(Path.Combine(tempDirectory, "sessions"));
            var settingsStore = new FileSettingsStore(
                Path.Combine(tempDirectory, "settings.json"),
                new DpapiSecretProtector());
            var orchestrator = new FakeAgentOrchestrator
            {
                ExceptionToThrow = new Exception("boom"),
            };

            var router = CreateRouter(
                sessionStore,
                settingsStore,
                new FakeExcelContextService(SelectionContext.Empty("No selection available.")),
                new FakeExcelCommandExecutor(),
                orchestrator);
            var responseJson = InvokeRoute(
                router,
                "{\"type\":\"bridge.runSkill\",\"requestId\":\"req-1\",\"payload\":{\"userInput\":\"/upload_data to ProjectA\",\"confirmed\":false}}");

            Assert.Contains("\"ok\":false", responseJson);
            Assert.Contains("\"code\":\"internal_error\"", responseJson);
        }

        public void Dispose()
        {
            if (Directory.Exists(tempDirectory))
            {
                Directory.Delete(tempDirectory, recursive: true);
            }
        }

        private static object CreateRouter(
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            string resolvedUiLocale = "en")
        {
            return CreateRouter(sessionStore, settingsStore, settings => resolvedUiLocale);
        }

        private static object CreateRouter(
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            Func<AppSettings, string> getResolvedUiLocale)
        {
            return CreateRouter(
                sessionStore,
                settingsStore,
                new FakeExcelContextService(SelectionContext.Empty("No selection available.")),
                new FakeExcelCommandExecutor(),
                new FakeAgentOrchestrator(),
                getResolvedUiLocale);
        }

        private static object CreateRouter(
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            IExcelContextService selectionContextService,
            IExcelCommandExecutor excelCommandExecutor,
            string resolvedUiLocale = "en")
        {
            return CreateRouter(sessionStore, settingsStore, selectionContextService, excelCommandExecutor, settings => resolvedUiLocale);
        }

        private static object CreateRouter(
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            IExcelContextService selectionContextService,
            IExcelCommandExecutor excelCommandExecutor,
            Func<AppSettings, string> getResolvedUiLocale)
        {
            return CreateRouter(
                sessionStore,
                settingsStore,
                selectionContextService,
                excelCommandExecutor,
                new FakeAgentOrchestrator(),
                getResolvedUiLocale);
        }

        private static object CreateRouter(
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            IExcelContextService selectionContextService,
            IExcelCommandExecutor excelCommandExecutor,
            IAgentOrchestrator agentOrchestrator,
            string resolvedUiLocale = "en")
        {
            return CreateRouter(sessionStore, settingsStore, selectionContextService, excelCommandExecutor, agentOrchestrator, settings => resolvedUiLocale);
        }

        private static object CreateRouter(
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            IExcelContextService selectionContextService,
            IExcelCommandExecutor excelCommandExecutor,
            IAgentOrchestrator agentOrchestrator,
            Func<AppSettings, string> getResolvedUiLocale)
        {
            var sharedCookies = new SharedCookieContainer();
            var cookieStore = new FileCookieStore(
                Path.Combine(Path.GetTempPath(), "OfficeAgent.Router.Tests", "cookies", Guid.NewGuid().ToString("N"), "cookies.json"),
                new DpapiSecretProtector());

            var addInAssembly = Assembly.LoadFrom(ResolveAddInAssemblyPath());
            var routerType = addInAssembly.GetType(
                "OfficeAgent.ExcelAddIn.WebBridge.WebMessageRouter",
                throwOnError: true);
            return Activator.CreateInstance(
                routerType,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                args: new object[] { sessionStore, settingsStore, selectionContextService, excelCommandExecutor, agentOrchestrator, sharedCookies, cookieStore, getResolvedUiLocale },
                culture: null);
        }

        private static string InvokeRoute(object router, string requestJson)
        {
            var routeMethod = router.GetType().GetMethod(
                "Route",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return (string)routeMethod.Invoke(router, new object[] { requestJson });
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

        private sealed class FakeExcelContextService : IExcelContextService
        {
            private readonly SelectionContext selectionContext;

            public FakeExcelContextService(SelectionContext selectionContext)
            {
                this.selectionContext = selectionContext;
            }

            public SelectionContext GetCurrentSelectionContext()
            {
                return selectionContext;
            }
        }

        private sealed class FakeExcelCommandExecutor : IExcelCommandExecutor
        {
            public int ExecuteCalls { get; private set; }

            public int PreviewCalls { get; private set; }

            public Exception ExecuteException { get; set; }

            public ExcelCommandResult ExecuteResult { get; set; } = new ExcelCommandResult
            {
                CommandType = ExcelCommandTypes.ReadSelectionTable,
                RequiresConfirmation = false,
                Status = "completed",
                Message = "Executed.",
            };

            public ExcelCommandResult PreviewResult { get; set; } = new ExcelCommandResult
            {
                CommandType = ExcelCommandTypes.AddWorksheet,
                RequiresConfirmation = true,
                Status = "preview",
                Message = "Preview ready.",
            };

            public ExcelCommandResult Preview(ExcelCommand command)
            {
                PreviewCalls++;
                return PreviewResult;
            }

            public ExcelCommandResult Execute(ExcelCommand command)
            {
                ExecuteCalls++;
                if (ExecuteException != null)
                {
                    throw ExecuteException;
                }

                return ExecuteResult;
            }
        }

        private sealed class FakeAgentOrchestrator : IAgentOrchestrator
        {
            public AgentCommandEnvelope LastEnvelope { get; private set; }

            public Exception ExceptionToThrow { get; set; }

            public AgentCommandResult Result { get; set; } = new AgentCommandResult
            {
                Route = AgentRouteTypes.Chat,
                Status = "completed",
                Message = "General chat routing is not implemented yet.",
            };

            public AgentCommandResult Execute(AgentCommandEnvelope envelope)
            {
                LastEnvelope = envelope;
                if (ExceptionToThrow != null)
                {
                    throw ExceptionToThrow;
                }

                return Result;
            }

            public System.Threading.Tasks.Task<AgentCommandResult> ExecuteAsync(AgentCommandEnvelope envelope)
            {
                return System.Threading.Tasks.Task.FromResult(Execute(envelope));
            }
        }
    }
}
