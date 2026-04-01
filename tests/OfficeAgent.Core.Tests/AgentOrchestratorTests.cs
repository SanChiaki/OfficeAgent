using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Orchestration;
using OfficeAgent.Core.Services;
using OfficeAgent.Core.Skills;
using Xunit;

namespace OfficeAgent.Core.Tests
{
    public sealed class AgentOrchestratorTests
    {
        private const string UploadToProjectA = "\u628A\u9009\u4E2D\u6570\u636E\u4E0A\u4F20\u5230\u9879\u76EEA";
        private const string ProjectA = "\u9879\u76EEA";
        private const string SummarizeRequest = "\u5E2E\u6211\u603B\u7ED3\u4E00\u4E0B\u8FD9\u4E2A\u5DE5\u4F5C\u7C3F";
        private const string UploadWithoutTargetProject = "upload project data";
        private const string UploadWithoutTargetSeparator = "\u628A\u9009\u4E2D\u6570\u636E\u4E0A\u4F20\u9879\u76EEA";

        [Fact]
        public void ExecuteRoutesSlashUploadCommandToTheUploadDataSkill()
        {
            var orchestrator = CreateOrchestrator();

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                UserInput = $"/upload_data {UploadToProjectA}",
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Skill, result.Route);
            Assert.Equal(SkillNames.UploadData, result.SkillName);
            Assert.True(result.RequiresConfirmation);
            Assert.Equal("preview", result.Status);
            Assert.NotNull(result.UploadPreview);
        }

        [Fact]
        public void ExecuteRoutesNaturalLanguageUploadIntentToTheUploadDataSkill()
        {
            var orchestrator = CreateOrchestrator();

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                UserInput = UploadToProjectA,
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Skill, result.Route);
            Assert.Equal(SkillNames.UploadData, result.SkillName);
            Assert.True(result.RequiresConfirmation);
            Assert.Equal(ProjectA, result.UploadPreview.ProjectName);
        }

        [Fact]
        public void ExecuteReturnsChatFallbackForUnknownUserInput()
        {
            var orchestrator = CreateOrchestrator();

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                UserInput = SummarizeRequest,
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Chat, result.Route);
            Assert.Equal("completed", result.Status);
            Assert.Contains("not implemented", result.Message, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ExecuteReturnsChatFallbackForEnglishUploadTextWithoutATargetProject()
        {
            var orchestrator = CreateOrchestrator();

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                UserInput = UploadWithoutTargetProject,
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Chat, result.Route);
            Assert.Equal("completed", result.Status);
            Assert.Contains("not implemented", result.Message, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ExecuteReturnsChatFallbackForChineseUploadTextWithoutExplicitTargetSeparator()
        {
            var orchestrator = CreateOrchestrator();

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                UserInput = UploadWithoutTargetSeparator,
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Chat, result.Route);
            Assert.Equal("completed", result.Status);
            Assert.Contains("not implemented", result.Message, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ExecuteUsesReadStepAndReturnsAFrozenPlanPreviewForAgentDispatch()
        {
            var plannerClient = new FakeLlmPlannerClient(
                PlannerJson.ReadStep(),
                PlannerJson.Plan());
            var excelCommandExecutor = new FakeExcelCommandExecutor();
            var orchestrator = CreateOrchestrator(
                plannerClient: plannerClient,
                excelCommandExecutor: excelCommandExecutor);

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                DispatchMode = AgentDispatchModes.Agent,
                SessionId = "session-1",
                UserInput = "Create a summary sheet from the current selection",
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Plan, result.Route);
            Assert.Equal("preview", result.Status);
            Assert.True(result.RequiresConfirmation);
            Assert.NotNull(result.Planner);
            Assert.Equal(PlannerResponseModes.Plan, result.Planner.Mode);
            Assert.Equal(2, plannerClient.Requests.Count);
            Assert.Single(plannerClient.Requests[1].Observations);
            Assert.Equal(1, excelCommandExecutor.ExecuteCalls);
            Assert.Equal(ExcelCommandTypes.ReadSelectionTable, excelCommandExecutor.LastExecutedCommand.CommandType);
        }

        [Fact]
        public void ExecuteReturnsControlledFailureWhenPlannerReturnsAnUnsupportedPlanStep()
        {
            var plannerClient = new FakeLlmPlannerClient(PlannerJson.InvalidPlan());
            var orchestrator = CreateOrchestrator(plannerClient: plannerClient);

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                DispatchMode = AgentDispatchModes.Agent,
                SessionId = "session-1",
                UserInput = "Do something unsupported",
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Chat, result.Route);
            Assert.Equal("failed", result.Status);
            Assert.False(result.RequiresConfirmation);
            Assert.Contains("supported", result.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ExecuteRunsTheFrozenPlanThroughThePlanExecutorWhenTheUserConfirms()
        {
            var plannerClient = new FakeLlmPlannerClient();
            var planExecutor = new FakePlanExecutor
            {
                Result = new PlanExecutionJournal
                {
                    HasFailures = true,
                    ErrorMessage = "Step 2 failed.",
                    Steps = new[]
                    {
                        new PlanExecutionJournalStep
                        {
                            Type = ExcelCommandTypes.AddWorksheet,
                            Title = "Add worksheet Summary",
                            Status = "completed",
                        },
                        new PlanExecutionJournalStep
                        {
                            Type = ExcelCommandTypes.WriteRange,
                            Title = "Write summary rows",
                            Status = "failed",
                            ErrorMessage = "Worksheet is protected.",
                        },
                        new PlanExecutionJournalStep
                        {
                            Type = SkillNames.UploadData,
                            Title = "Upload selection",
                            Status = "skipped",
                        },
                    },
                },
            };
            var frozenPlan = new AgentPlan
            {
                Summary = "Create a summary sheet and upload the selection",
                Steps = new[]
                {
                    new AgentPlanStep
                    {
                        Type = ExcelCommandTypes.AddWorksheet,
                        Args = JObject.FromObject(new
                        {
                            newSheetName = "Summary",
                        }),
                    },
                    new AgentPlanStep
                    {
                        Type = ExcelCommandTypes.WriteRange,
                        Args = JObject.FromObject(new
                        {
                            targetAddress = "Summary!A1:B2",
                            values = new[]
                            {
                                new[] { "Name", "Region" },
                            },
                        }),
                    },
                    new AgentPlanStep
                    {
                        Type = PlannerStepTypes.UploadData,
                        Args = JObject.FromObject(new
                        {
                            userInput = "把选中数据上传到项目A",
                        }),
                    },
                },
            };
            var orchestrator = CreateOrchestrator(
                plannerClient: plannerClient,
                planExecutor: planExecutor);

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                DispatchMode = AgentDispatchModes.Agent,
                SessionId = "session-1",
                UserInput = "Create a summary sheet and upload the selection",
                Confirmed = true,
                Plan = frozenPlan,
            });

            Assert.Equal(AgentRouteTypes.Plan, result.Route);
            Assert.Equal("failed", result.Status);
            Assert.False(result.RequiresConfirmation);
            Assert.Equal("Step 2 failed.", result.Message);
            Assert.Same(frozenPlan, planExecutor.LastPlan);
            Assert.Null(result.Planner);
            Assert.NotNull(result.Journal);
            Assert.Equal("completed", result.Journal.Steps[0].Status);
            Assert.Equal("failed", result.Journal.Steps[1].Status);
            Assert.Equal("skipped", result.Journal.Steps[2].Status);
            Assert.Empty(plannerClient.Requests);
        }

        [Fact]
        public void ExecuteUsesReadRangeStepToReadASpecificRange()
        {
            var plannerClient = new FakeLlmPlannerClient(
                PlannerJson.ReadRangeStep("A1:D10", "Sheet2"),
                PlannerJson.Plan());
            var excelCommandExecutor = new FakeExcelCommandExecutor();
            var orchestrator = CreateOrchestrator(
                plannerClient: plannerClient,
                excelCommandExecutor: excelCommandExecutor);

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                DispatchMode = AgentDispatchModes.Agent,
                SessionId = "session-1",
                UserInput = "Read Sheet2 A1:D10 and create a summary",
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Plan, result.Route);
            Assert.Equal("preview", result.Status);
            Assert.True(result.RequiresConfirmation);
            Assert.Equal(2, plannerClient.Requests.Count);
            Assert.Single(plannerClient.Requests[1].Observations);
            Assert.Equal("excel.table", plannerClient.Requests[1].Observations[0].Kind);
            Assert.Equal(1, excelCommandExecutor.ExecuteCalls);
            Assert.Equal(ExcelCommandTypes.ReadRange, excelCommandExecutor.LastExecutedCommand.CommandType);
            Assert.Equal("Sheet2", excelCommandExecutor.LastExecutedCommand.SheetName);
            Assert.Equal("A1:D10", excelCommandExecutor.LastExecutedCommand.TargetAddress);
        }

        [Fact]
        public void ExecuteUsesFetchUrlStepToRetrieveExternalData()
        {
            var fetchClient = new FakeAgentFetchClient();
            var plannerClient = new FakeLlmPlannerClient(
                PlannerJson.FetchUrlStep("http://localhost:3200/api/performance"),
                PlannerJson.Plan());
            var orchestrator = CreateOrchestrator(
                plannerClient: plannerClient,
                fetchClient: fetchClient);

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                DispatchMode = AgentDispatchModes.Agent,
                SessionId = "session-1",
                UserInput = "查询所有项目业绩并写入Excel",
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Plan, result.Route);
            Assert.Equal("preview", result.Status);
            Assert.Equal(1, fetchClient.FetchCalls);
            Assert.Equal("http://localhost:3200/api/performance", fetchClient.LastUrl);
            Assert.Equal(2, plannerClient.Requests.Count);
            Assert.Single(plannerClient.Requests[1].Observations);
            Assert.Equal("fetch.response", plannerClient.Requests[1].Observations[0].Kind);
            Assert.NotNull(plannerClient.Requests[1].Observations[0].Data);
        }

        [Fact]
        public void ExecuteIncludesFetchErrorAsObservationAndPlannerRecoverable()
        {
            var fetchClient = new FakeAgentFetchClient
            {
                Result = new FetchResult
                {
                    Success = false,
                    StatusCode = 500,
                    ErrorMessage = "Internal Server Error",
                },
            };
            var plannerClient = new FakeLlmPlannerClient(
                PlannerJson.FetchUrlStep("http://localhost:3200/api/error"),
                PlannerJson.Message("The API returned an error. I cannot retrieve the data right now."));
            var orchestrator = CreateOrchestrator(
                plannerClient: plannerClient,
                fetchClient: fetchClient);

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                DispatchMode = AgentDispatchModes.Agent,
                SessionId = "session-1",
                UserInput = "查询不存在的接口",
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Chat, result.Route);
            Assert.Equal("completed", result.Status);
            Assert.Equal(2, plannerClient.Requests.Count);
            Assert.Single(plannerClient.Requests[1].Observations);
            Assert.Equal("fetch.error", plannerClient.Requests[1].Observations[0].Kind);
        }

        [Fact]
        public void ExecuteReturnsFailureWhenFetchUrlStepIsUsedButFetchClientIsNull()
        {
            var plannerClient = new FakeLlmPlannerClient(
                PlannerJson.FetchUrlStep("http://localhost:3200/api/performance"));
            var orchestrator = CreateOrchestrator(
                plannerClient: plannerClient,
                fetchClient: null);

            var result = orchestrator.Execute(new AgentCommandEnvelope
            {
                DispatchMode = AgentDispatchModes.Agent,
                SessionId = "session-1",
                UserInput = "查询业绩",
                Confirmed = false,
            });

            Assert.Equal(AgentRouteTypes.Chat, result.Route);
            Assert.Equal("failed", result.Status);
            Assert.Contains("not supported or not configured", result.Message, StringComparison.OrdinalIgnoreCase);
        }

        private static AgentOrchestrator CreateOrchestrator()
        {
            return CreateOrchestrator(
                plannerClient: new FakeLlmPlannerClient(),
                excelCommandExecutor: new FakeExcelCommandExecutor(),
                planExecutor: new FakePlanExecutor());
        }

        private static AgentOrchestrator CreateOrchestrator(
            ILlmPlannerClient plannerClient,
            FakeExcelCommandExecutor excelCommandExecutor = null,
            FakePlanExecutor planExecutor = null,
            IAgentFetchClient fetchClient = null)
        {
            excelCommandExecutor = excelCommandExecutor ?? new FakeExcelCommandExecutor();
            var skill = new UploadDataSkill(
                excelCommandExecutor,
                new FakeUploadDataGateway());

            return new AgentOrchestrator(
                new SkillRegistry(skill),
                new FakeExcelContextService(),
                excelCommandExecutor,
                plannerClient,
                planExecutor ?? new FakePlanExecutor(),
                fetchClient,
                () => new AppSettings { BaseUrl = "http://localhost:3200" });
        }

        private sealed class FakeExcelCommandExecutor : IExcelCommandExecutor
        {
            public ExcelCommand LastExecutedCommand { get; private set; }

            public int ExecuteCalls { get; private set; }

            public ExcelCommandResult Preview(ExcelCommand command)
            {
                throw new System.NotSupportedException();
            }

            public ExcelCommandResult Execute(ExcelCommand command)
            {
                ExecuteCalls++;
                LastExecutedCommand = command;

                if (string.Equals(command.CommandType, ExcelCommandTypes.ReadRange, System.StringComparison.Ordinal))
                {
                    var sheetName = string.IsNullOrWhiteSpace(command.SheetName) ? "Sheet1" : command.SheetName;
                    return new ExcelCommandResult
                    {
                        CommandType = ExcelCommandTypes.ReadRange,
                        RequiresConfirmation = false,
                        Status = "completed",
                        Message = $"Read range from {sheetName} {command.TargetAddress}.",
                        Table = new ExcelTableData
                        {
                            SheetName = sheetName,
                            Address = command.TargetAddress,
                            Headers = new[] { "Name", "Region" },
                            Rows = new[]
                            {
                                new[] { "Project A", "CN" },
                                new[] { "Project B", "US" },
                            },
                        },
                    };
                }

                return new ExcelCommandResult
                {
                    CommandType = ExcelCommandTypes.ReadSelectionTable,
                    RequiresConfirmation = false,
                    Status = "completed",
                    Message = "Read selection from Sheet1 A1:C3.",
                    Table = new ExcelTableData
                    {
                        SheetName = "Sheet1",
                        Address = "A1:C3",
                        Headers = new[] { "Name", "Region" },
                        Rows = new[]
                        {
                            new[] { "Project A", "CN" },
                            new[] { "Project B", "US" },
                        },
                    },
                };
            }
        }

        private sealed class FakeExcelContextService : IExcelContextService
        {
            public SelectionContext GetCurrentSelectionContext()
            {
                return new SelectionContext
                {
                    HasSelection = true,
                    WorkbookName = "Quarterly Report.xlsx",
                    SheetName = "Sheet1",
                    Address = "A1:C3",
                    RowCount = 3,
                    ColumnCount = 2,
                    IsContiguous = true,
                    HeaderPreview = new[] { "Name", "Region" },
                    SampleRows = new[]
                    {
                        new[] { "Project A", "CN" },
                        new[] { "Project B", "US" },
                    },
                };
            }
        }

        private sealed class FakeUploadDataGateway : IUploadDataGateway
        {
            public UploadExecutionResult Upload(UploadPreview preview)
            {
                return new UploadExecutionResult
                {
                    SavedCount = preview?.Records?.Length ?? 0,
                    Message = $"Uploaded {preview?.Records?.Length ?? 0} row(s).",
                };
            }
        }

        private sealed class FakeLlmPlannerClient : ILlmPlannerClient
        {
            private readonly Queue<string> responses;

            public FakeLlmPlannerClient(params string[] responses)
            {
                this.responses = new Queue<string>(responses ?? Array.Empty<string>());
            }

            public List<PlannerRequest> Requests { get; } = new List<PlannerRequest>();

            public string Complete(PlannerRequest request)
            {
                Requests.Add(request);
                return responses.Count > 0
                    ? responses.Dequeue()
                    : PlannerJson.Message("General chat routing is not implemented yet.");
            }

            public System.Threading.Tasks.Task<string> CompleteAsync(PlannerRequest request)
            {
                return System.Threading.Tasks.Task.FromResult(Complete(request));
            }
        }

        private sealed class FakePlanExecutor : IPlanExecutor
        {
            public AgentPlan LastPlan { get; private set; }

            public PlanExecutionJournal Result { get; set; } = new PlanExecutionJournal
            {
                Steps = Array.Empty<PlanExecutionJournalStep>(),
            };

            public PlanExecutionJournal Execute(AgentPlan plan)
            {
                LastPlan = plan;
                return Result;
            }
        }

        private static class PlannerJson
        {
            public static string Message(string assistantMessage)
            {
                return "{"
                    + "\"mode\":\"message\","
                    + $"\"assistantMessage\":\"{assistantMessage}\""
                    + "}";
            }

            public static string ReadStep()
            {
                return "{"
                    + "\"mode\":\"read_step\","
                    + "\"assistantMessage\":\"I need the full selection before I can write a plan.\","
                    + "\"step\":{"
                    + "\"type\":\"excel.readSelectionTable\","
                    + "\"args\":{}"
                    + "}"
                    + "}";
            }

            public static string Plan()
            {
                return "{"
                    + "\"mode\":\"plan\","
                    + "\"assistantMessage\":\"I prepared a plan. Review it before Excel is changed.\","
                    + "\"plan\":{"
                    + "\"summary\":\"Create a Summary sheet and write the selected rows.\","
                    + "\"steps\":["
                    + "{"
                    + "\"type\":\"excel.addWorksheet\","
                    + "\"args\":{\"newSheetName\":\"Summary\"}"
                    + "},"
                    + "{"
                    + "\"type\":\"excel.writeRange\","
                    + "\"args\":{\"targetAddress\":\"Summary!A1:B3\",\"values\":[[\"Name\",\"Region\"],[\"Project A\",\"CN\"],[\"Project B\",\"US\"]]}"
                    + "}"
                    + "]"
                    + "}"
                    + "}";
            }

            public static string InvalidPlan()
            {
                return "{"
                    + "\"mode\":\"plan\","
                    + "\"assistantMessage\":\"I prepared a plan.\","
                    + "\"plan\":{"
                    + "\"summary\":\"Do something unsupported.\","
                    + "\"steps\":["
                    + "{"
                    + "\"type\":\"excel.formatCells\","
                    + "\"args\":{}"
                    + "}"
                    + "]"
                    + "}"
                    + "}";
            }

            public static string ReadRangeStep(string address, string sheetName = null)
            {
                var args = string.IsNullOrWhiteSpace(sheetName)
                    ? $"\"address\":\"{address}\""
                    : $"\"address\":\"{address}\",\"sheetName\":\"{sheetName}\"";
                return "{"
                    + "\"mode\":\"read_step\","
                    + "\"assistantMessage\":\"I need to read a specific range.\","
                    + "\"step\":{"
                    + "\"type\":\"excel.readRange\","
                    + $"\"args\":{{{args}}}"
                    + "}"
                    + "}";
            }

            public static string FetchUrlStep(string url)
            {
                return "{"
                    + "\"mode\":\"read_step\","
                    + "\"assistantMessage\":\"I need to fetch data from an API.\","
                    + "\"step\":{"
                    + "\"type\":\"fetch.url\","
                    + $"\"args\":{{\"url\":\"{url}\"}}"
                    + "}"
                    + "}";
            }
        }

        private sealed class FakeAgentFetchClient : IAgentFetchClient
        {
            public string LastUrl { get; private set; }

            public int FetchCalls { get; private set; }

            public FetchResult Result { get; set; } = new FetchResult
            {
                Success = true,
                StatusCode = 200,
                Body = "{\"data\":[{\"name\":\"Project A\"}]}",
            };

            public Task<FetchResult> FetchAsync(string url)
            {
                FetchCalls++;
                LastUrl = url;
                return Task.FromResult(Result);
            }
        }
    }
}
