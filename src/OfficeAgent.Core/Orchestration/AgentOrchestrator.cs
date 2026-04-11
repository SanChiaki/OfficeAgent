using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;

namespace OfficeAgent.Core.Orchestration
{
    public sealed class AgentOrchestrator : IAgentOrchestrator
    {
        private const int MaxPlannerAttempts = 5;

        private readonly ISkillRegistry skillRegistry;
        private readonly IExcelContextService excelContextService;
        private readonly IExcelCommandExecutor excelCommandExecutor;
        private readonly ILlmPlannerClient plannerClient;
        private readonly IPlanExecutor planExecutor;
        private readonly IAgentFetchClient fetchClient;
        private readonly Func<AppSettings> loadSettings;

        public AgentOrchestrator(ISkillRegistry skillRegistry)
            : this(skillRegistry, null, null, null, null, null, null)
        {
        }

        public AgentOrchestrator(
            ISkillRegistry skillRegistry,
            IExcelContextService excelContextService,
            IExcelCommandExecutor excelCommandExecutor,
            ILlmPlannerClient plannerClient,
            IPlanExecutor planExecutor,
            IAgentFetchClient fetchClient = null,
            Func<AppSettings> loadSettings = null)
        {
            this.skillRegistry = skillRegistry ?? throw new ArgumentNullException(nameof(skillRegistry));
            this.excelContextService = excelContextService;
            this.excelCommandExecutor = excelCommandExecutor;
            this.plannerClient = plannerClient;
            this.planExecutor = planExecutor;
            this.fetchClient = fetchClient;
            this.loadSettings = loadSettings;
        }

        public AgentCommandResult Execute(AgentCommandEnvelope envelope)
        {
            if (envelope == null)
            {
                throw new ArgumentNullException(nameof(envelope));
            }

            if (string.Equals(envelope.DispatchMode, AgentDispatchModes.Agent, StringComparison.Ordinal))
            {
                return ExecuteAgentFlow(envelope);
            }

            var skillName = ResolveSkillName(envelope);
            if (!string.IsNullOrWhiteSpace(skillName))
            {
                envelope.SkillName = skillName;
                OfficeAgentLog.Info("agent", "route.skill", $"Routing to skill {skillName}.");
                return skillRegistry.Resolve(skillName).Execute(envelope);
            }

            OfficeAgentLog.Info("agent", "route.chat", "Falling back to chat route.");
            return new AgentCommandResult
            {
                Route = AgentRouteTypes.Chat,
                Status = "completed",
                Message = "General chat routing is not implemented yet. Use /upload_data ... or a direct Excel command.",
            };
        }

        public async Task<AgentCommandResult> ExecuteAsync(AgentCommandEnvelope envelope)
        {
            if (envelope == null)
            {
                throw new ArgumentNullException(nameof(envelope));
            }

            if (string.Equals(envelope.DispatchMode, AgentDispatchModes.Agent, StringComparison.Ordinal))
            {
                return await ExecuteAgentFlowAsync(envelope).ConfigureAwait(true);
            }

            var skillName = ResolveSkillName(envelope);
            if (!string.IsNullOrWhiteSpace(skillName))
            {
                envelope.SkillName = skillName;
                OfficeAgentLog.Info("agent", "route.skill", $"Routing to skill {skillName}.");
                return skillRegistry.Resolve(skillName).Execute(envelope);
            }

            OfficeAgentLog.Info("agent", "route.chat", "Falling back to chat route.");
            return new AgentCommandResult
            {
                Route = AgentRouteTypes.Chat,
                Status = "completed",
                Message = "General chat routing is not implemented yet. Use /upload_data ... or a direct Excel command.",
            };
        }

        private AgentCommandResult ExecuteAgentFlow(AgentCommandEnvelope envelope)
        {
            if (envelope.Confirmed && envelope.Plan != null)
            {
                return ExecuteFrozenPlan(envelope.Plan);
            }

            if (plannerClient == null || excelContextService == null || excelCommandExecutor == null)
            {
                return CreatePlannerFailure("Agent planning is not configured for this OfficeAgent host.");
            }

            var request = BuildPlannerRequest(envelope);

            for (var attempt = 0; attempt < MaxPlannerAttempts; attempt++)
            {
                var plannerResponse = DeserializePlannerResponse(plannerClient.Complete(request));
                var validationError = ValidatePlannerResponse(plannerResponse);
                if (!string.IsNullOrWhiteSpace(validationError))
                {
                    OfficeAgentLog.Warn("agent", "planner.invalid", validationError);
                    return CreatePlannerFailure(validationError);
                }

                if (string.Equals(plannerResponse.Mode, PlannerResponseModes.Message, StringComparison.Ordinal))
                {
                    return new AgentCommandResult
                    {
                        Route = AgentRouteTypes.Chat,
                        Status = "completed",
                        Message = plannerResponse.AssistantMessage,
                        Planner = plannerResponse,
                    };
                }

                if (string.Equals(plannerResponse.Mode, PlannerResponseModes.ReadStep, StringComparison.Ordinal))
                {
                    var observation = ExecuteReadStep(plannerResponse.Step);
                    if (observation == null)
                    {
                        return CreatePlannerFailure($"The read step type '{plannerResponse.Step?.Type}' is not supported or not configured.");
                    }

                    if (string.Equals(observation.Kind, "fetch.error", StringComparison.Ordinal))
                    {
                        return new AgentCommandResult
                        {
                            Route = AgentRouteTypes.Chat,
                            Status = "failed",
                            RequiresConfirmation = false,
                            Message = observation.Message,
                        };
                    }

                    request.Observations = request.Observations
                        .Concat(new[] { observation })
                        .ToArray();
                    continue;
                }

                if (string.Equals(plannerResponse.Mode, PlannerResponseModes.Plan, StringComparison.Ordinal))
                {
                    return new AgentCommandResult
                    {
                        Route = AgentRouteTypes.Plan,
                        Status = "preview",
                        RequiresConfirmation = true,
                        Message = plannerResponse.AssistantMessage,
                        Planner = plannerResponse,
                    };
                }
            }

            return CreatePlannerFailure("I could not produce a safe plan. Rephrase the request or use an explicit command.");
        }

        private async Task<AgentCommandResult> ExecuteAgentFlowAsync(AgentCommandEnvelope envelope)
        {
            if (envelope.Confirmed && envelope.Plan != null)
            {
                return ExecuteFrozenPlan(envelope.Plan);
            }

            if (plannerClient == null || excelContextService == null || excelCommandExecutor == null)
            {
                return CreatePlannerFailure("Agent planning is not configured for this OfficeAgent host.");
            }

            // GetCurrentSelectionContext is a COM call — runs on the calling (UI) thread before first await
            var request = BuildPlannerRequest(envelope);

            for (var attempt = 0; attempt < MaxPlannerAttempts; attempt++)
            {
                // ConfigureAwait(true) keeps the continuation on the UI thread (SynchronizationContext),
                // so COM calls after the await are safe.
                var rawResponse = await plannerClient.CompleteAsync(request).ConfigureAwait(true);
                var plannerResponse = DeserializePlannerResponse(rawResponse);
                var validationError = ValidatePlannerResponse(plannerResponse);
                if (!string.IsNullOrWhiteSpace(validationError))
                {
                    OfficeAgentLog.Warn("agent", "planner.invalid", validationError);
                    return CreatePlannerFailure(validationError);
                }

                if (string.Equals(plannerResponse.Mode, PlannerResponseModes.Message, StringComparison.Ordinal))
                {
                    return new AgentCommandResult
                    {
                        Route = AgentRouteTypes.Chat,
                        Status = "completed",
                        Message = plannerResponse.AssistantMessage,
                        Planner = plannerResponse,
                    };
                }

                if (string.Equals(plannerResponse.Mode, PlannerResponseModes.ReadStep, StringComparison.Ordinal))
                {
                    // Back on the UI thread here — COM call is safe
                    var observation = await ExecuteReadStepAsync(plannerResponse.Step).ConfigureAwait(true);
                    if (observation == null)
                    {
                        return CreatePlannerFailure($"The read step type '{plannerResponse.Step?.Type}' is not supported or not configured.");
                    }

                    if (string.Equals(observation.Kind, "fetch.error", StringComparison.Ordinal))
                    {
                        return new AgentCommandResult
                        {
                            Route = AgentRouteTypes.Chat,
                            Status = "failed",
                            RequiresConfirmation = false,
                            Message = observation.Message,
                        };
                    }

                    request.Observations = request.Observations
                        .Concat(new[] { observation })
                        .ToArray();
                    continue;
                }

                if (string.Equals(plannerResponse.Mode, PlannerResponseModes.Plan, StringComparison.Ordinal))
                {
                    return new AgentCommandResult
                    {
                        Route = AgentRouteTypes.Plan,
                        Status = "preview",
                        RequiresConfirmation = true,
                        Message = plannerResponse.AssistantMessage,
                        Planner = plannerResponse,
                    };
                }
            }

            return CreatePlannerFailure("I could not produce a safe plan. Rephrase the request or use an explicit command.");
        }

        private PlannerRequest BuildPlannerRequest(AgentCommandEnvelope envelope)
        {
            var settings = loadSettings?.Invoke();
            var apiBaseUrl = settings != null ? AppSettings.NormalizeOptionalUrl(settings.BusinessBaseUrl) : string.Empty;

            return new PlannerRequest
            {
                SessionId = envelope.SessionId ?? string.Empty,
                UserInput = envelope.UserInput?.Trim() ?? string.Empty,
                SelectionContext = excelContextService.GetCurrentSelectionContext(),
                ConversationHistory = envelope.ConversationHistory ?? System.Array.Empty<ConversationTurn>(),
                ApiBaseUrl = apiBaseUrl,
            };
        }

        private PlannerObservation ExecuteReadStep(PlannerStep step)
        {
            if (string.Equals(step.Type, PlannerStepTypes.ReadSelectionTable, StringComparison.Ordinal))
            {
                var readResult = excelCommandExecutor.Execute(new ExcelCommand
                {
                    CommandType = ExcelCommandTypes.ReadSelectionTable,
                    Confirmed = false,
                });

                return new PlannerObservation
                {
                    Kind = "excel.table",
                    Message = readResult.Message,
                    Table = readResult.Table,
                };
            }

            if (string.Equals(step.Type, PlannerStepTypes.ReadRange, StringComparison.Ordinal))
            {
                var address = step.Args?["address"]?.Value<string>() ?? string.Empty;
                var sheetName = step.Args?["sheetName"]?.Value<string>() ?? string.Empty;

                var readResult = excelCommandExecutor.Execute(new ExcelCommand
                {
                    CommandType = ExcelCommandTypes.ReadRange,
                    SheetName = sheetName,
                    TargetAddress = address,
                    Confirmed = false,
                });

                return new PlannerObservation
                {
                    Kind = "excel.table",
                    Message = readResult.Message,
                    Table = readResult.Table,
                };
            }

            if (string.Equals(step.Type, PlannerStepTypes.FetchUrl, StringComparison.Ordinal))
            {
                return ExecuteFetchReadStep(step);
            }

            return null;
        }

        private async Task<PlannerObservation> ExecuteReadStepAsync(PlannerStep step)
        {
            if (string.Equals(step.Type, PlannerStepTypes.ReadSelectionTable, StringComparison.Ordinal))
            {
                // Back on UI thread — COM call is safe
                var readResult = excelCommandExecutor.Execute(new ExcelCommand
                {
                    CommandType = ExcelCommandTypes.ReadSelectionTable,
                    Confirmed = false,
                });

                return new PlannerObservation
                {
                    Kind = "excel.table",
                    Message = readResult.Message,
                    Table = readResult.Table,
                };
            }

            if (string.Equals(step.Type, PlannerStepTypes.ReadRange, StringComparison.Ordinal))
            {
                // Back on UI thread — COM call is safe
                var address = step.Args?["address"]?.Value<string>() ?? string.Empty;
                var sheetName = step.Args?["sheetName"]?.Value<string>() ?? string.Empty;

                var readResult = excelCommandExecutor.Execute(new ExcelCommand
                {
                    CommandType = ExcelCommandTypes.ReadRange,
                    SheetName = sheetName,
                    TargetAddress = address,
                    Confirmed = false,
                });

                return new PlannerObservation
                {
                    Kind = "excel.table",
                    Message = readResult.Message,
                    Table = readResult.Table,
                };
            }

            if (string.Equals(step.Type, PlannerStepTypes.FetchUrl, StringComparison.Ordinal))
            {
                return await ExecuteFetchReadStepAsync(step).ConfigureAwait(true);
            }

            return null;
        }

        private PlannerObservation ExecuteFetchReadStep(PlannerStep step)
        {
            if (fetchClient == null)
            {
                OfficeAgentLog.Warn("agent", "fetch.not_configured", "Fetch client is not configured.");
                return null;
            }

            var url = step.Args?["url"]?.Value<string>() ?? string.Empty;
            var headers = step.Args?["headers"]?.Type == JTokenType.Object ? step.Args["headers"] as JObject : null;
            OfficeAgentLog.Info("agent", "fetch.begin", $"Fetching URL: {url}");

            var fetchResult = fetchClient.FetchAsync(url, headers).GetAwaiter().GetResult();

            if (!fetchResult.Success)
            {
                OfficeAgentLog.Warn("agent", "fetch.failed", $"Fetch failed for {url}: {fetchResult.ErrorMessage}");
                return new PlannerObservation
                {
                    Kind = "fetch.error",
                    Message = fetchResult.ErrorMessage,
                };
            }

            OfficeAgentLog.Info("agent", "fetch.completed", $"Fetched {url} — {(int)fetchResult.StatusCode}");

            JToken data = null;
            try
            {
                data = JToken.Parse(fetchResult.Body);
            }
            catch (JsonException)
            {
                // Body is not JSON — store as string in a wrapper object
                data = new JObject { ["text"] = fetchResult.Body };
            }

            return new PlannerObservation
            {
                Kind = "fetch.response",
                Message = $"Fetched {url} — {(int)fetchResult.StatusCode}",
                Data = data,
            };
        }

        private async Task<PlannerObservation> ExecuteFetchReadStepAsync(PlannerStep step)
        {
            if (fetchClient == null)
            {
                OfficeAgentLog.Warn("agent", "fetch.not_configured", "Fetch client is not configured.");
                return null;
            }

            var url = step.Args?["url"]?.Value<string>() ?? string.Empty;
            var headers = step.Args?["headers"]?.Type == JTokenType.Object ? step.Args["headers"] as JObject : null;
            OfficeAgentLog.Info("agent", "fetch.begin", $"Fetching URL: {url}");

            var fetchResult = await fetchClient.FetchAsync(url, headers).ConfigureAwait(false);

            if (!fetchResult.Success)
            {
                OfficeAgentLog.Warn("agent", "fetch.failed", $"Fetch failed for {url}: {fetchResult.ErrorMessage}");
                return new PlannerObservation
                {
                    Kind = "fetch.error",
                    Message = fetchResult.ErrorMessage,
                };
            }

            OfficeAgentLog.Info("agent", "fetch.completed", $"Fetched {url} — {(int)fetchResult.StatusCode}");

            JToken data = null;
            try
            {
                data = JToken.Parse(fetchResult.Body);
            }
            catch (JsonException)
            {
                // Body is not JSON — store as string in a wrapper object
                data = new JObject { ["text"] = fetchResult.Body };
            }

            return new PlannerObservation
            {
                Kind = "fetch.response",
                Message = $"Fetched {url} — {(int)fetchResult.StatusCode}",
                Data = data,
            };
        }

        private AgentCommandResult ExecuteFrozenPlan(AgentPlan plan)
        {
            if (plan == null)
            {
                return CreatePlannerFailure("A confirmed plan payload is required before execution can start.");
            }

            if (planExecutor == null)
            {
                return CreatePlannerFailure("Agent plan execution is not configured for this OfficeAgent host.");
            }

            var journal = planExecutor.Execute(plan) ?? new PlanExecutionJournal();
            return new AgentCommandResult
            {
                Route = AgentRouteTypes.Plan,
                Status = journal.HasFailures ? "failed" : "completed",
                RequiresConfirmation = false,
                Message = journal.HasFailures
                    ? journal.ErrorMessage ?? "Plan execution failed."
                    : "Plan executed successfully.",
                Journal = journal,
            };
        }

        private static PlannerResponse DeserializePlannerResponse(string rawPlannerResponse)
        {
            try
            {
                return JsonConvert.DeserializeObject<PlannerResponse>(rawPlannerResponse ?? string.Empty) ?? new PlannerResponse();
            }
            catch (JsonException)
            {
                return new PlannerResponse();
            }
        }

        private static string ValidatePlannerResponse(PlannerResponse plannerResponse)
        {
            if (plannerResponse == null)
            {
                return "The planner returned an empty response.";
            }

            if (string.IsNullOrWhiteSpace(plannerResponse.AssistantMessage))
            {
                return "The planner response must include an assistantMessage.";
            }

            if (string.Equals(plannerResponse.Mode, PlannerResponseModes.Message, StringComparison.Ordinal))
            {
                return string.Empty;
            }

            if (string.Equals(plannerResponse.Mode, PlannerResponseModes.ReadStep, StringComparison.Ordinal))
            {
                if (plannerResponse.Step == null)
                {
                    return "The planner read_step must include a step object.";
                }

                if (string.Equals(plannerResponse.Step.Type, PlannerStepTypes.ReadSelectionTable, StringComparison.Ordinal))
                {
                    if (plannerResponse.Step.Args != null && plannerResponse.Step.Args.HasValues)
                    {
                        return "The planner read_step excel.readSelectionTable must have empty args.";
                    }

                    return string.Empty;
                }

                if (string.Equals(plannerResponse.Step.Type, PlannerStepTypes.ReadRange, StringComparison.Ordinal))
                {
                    var address = plannerResponse.Step.Args?["address"]?.Value<string>();
                    if (string.IsNullOrWhiteSpace(address))
                    {
                        return "The planner read_step excel.readRange requires an address arg.";
                    }

                    return string.Empty;
                }

                if (string.Equals(plannerResponse.Step.Type, PlannerStepTypes.FetchUrl, StringComparison.Ordinal))
                {
                    var url = plannerResponse.Step.Args?["url"]?.Value<string>();
                    if (string.IsNullOrWhiteSpace(url))
                    {
                        return "The planner read_step fetch.url requires a url arg.";
                    }

                    return string.Empty;
                }

                return "The planner can only use the supported read step types: excel.readSelectionTable, excel.readRange, fetch.url.";
            }

            if (string.Equals(plannerResponse.Mode, PlannerResponseModes.Plan, StringComparison.Ordinal))
            {
                if (plannerResponse.Plan == null ||
                    string.IsNullOrWhiteSpace(plannerResponse.Plan.Summary) ||
                    plannerResponse.Plan.Steps == null ||
                    plannerResponse.Plan.Steps.Length == 0)
                {
                    return "The planner must return a non-empty plan before Excel can be modified.";
                }

                foreach (var step in plannerResponse.Plan.Steps)
                {
                    if (!IsSupportedPlanStep(step?.Type))
                    {
                        return "The planner can only use supported write actions and skills.";
                    }
                }

                return string.Empty;
            }

            return "The planner returned an unsupported mode.";
        }

        private static bool IsSupportedPlanStep(string stepType)
        {
            return string.Equals(stepType, ExcelCommandTypes.WriteRange, StringComparison.Ordinal) ||
                   string.Equals(stepType, ExcelCommandTypes.AddWorksheet, StringComparison.Ordinal) ||
                   string.Equals(stepType, ExcelCommandTypes.RenameWorksheet, StringComparison.Ordinal) ||
                   string.Equals(stepType, ExcelCommandTypes.DeleteWorksheet, StringComparison.Ordinal) ||
                   string.Equals(stepType, PlannerStepTypes.UploadData, StringComparison.Ordinal);
        }

        private static AgentCommandResult CreatePlannerFailure(string message)
        {
            return new AgentCommandResult
            {
                Route = AgentRouteTypes.Chat,
                Status = "failed",
                RequiresConfirmation = false,
                Message = message,
            };
        }

        private static string ResolveSkillName(AgentCommandEnvelope envelope)
        {
            if (!string.IsNullOrWhiteSpace(envelope.SkillName))
            {
                return envelope.SkillName.Trim();
            }

            var input = envelope.UserInput?.Trim() ?? string.Empty;
            if (input.Length == 0)
            {
                return string.Empty;
            }

            if (input.StartsWith("/upload_data", StringComparison.OrdinalIgnoreCase))
            {
                return SkillNames.UploadData;
            }

            if (input.IndexOf("\u4E0A\u4F20\u5230", StringComparison.Ordinal) >= 0)
            {
                return SkillNames.UploadData;
            }

            if (Regex.IsMatch(input, "(?i)\\bupload\\b.+\\bto\\s+.+$"))
            {
                return SkillNames.UploadData;
            }

            return string.Empty;
        }
    }
}
