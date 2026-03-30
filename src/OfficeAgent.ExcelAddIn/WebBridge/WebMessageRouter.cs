using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Infrastructure.Storage;

namespace OfficeAgent.ExcelAddIn.WebBridge
{
    internal sealed class WebMessageRouter
    {
        private static readonly JsonSerializerSettings SerializerSettings = new JsonSerializerSettings
        {
            ContractResolver = new CamelCasePropertyNamesContractResolver(),
            NullValueHandling = NullValueHandling.Ignore,
        };

        private readonly IExcelContextService excelContextService;
        private readonly IExcelCommandExecutor excelCommandExecutor;
        private readonly IAgentOrchestrator agentOrchestrator;
        private readonly ConfirmationService confirmationService = new ConfirmationService();
        private readonly HashSet<string> allowedTypes = new HashSet<string>(StringComparer.Ordinal)
        {
            BridgeMessageTypes.Ping,
            BridgeMessageTypes.GetSettings,
            BridgeMessageTypes.GetSelectionContext,
            BridgeMessageTypes.GetSessions,
            BridgeMessageTypes.SaveSettings,
            BridgeMessageTypes.ExecuteExcelCommand,
            BridgeMessageTypes.RunSkill,
            BridgeMessageTypes.RunAgent,
        };
        private readonly FileSessionStore sessionStore;
        private readonly FileSettingsStore settingsStore;

        public WebMessageRouter(
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            IExcelContextService excelContextService,
            IExcelCommandExecutor excelCommandExecutor,
            IAgentOrchestrator agentOrchestrator)
        {
            this.sessionStore = sessionStore ?? throw new ArgumentNullException(nameof(sessionStore));
            this.settingsStore = settingsStore ?? throw new ArgumentNullException(nameof(settingsStore));
            this.excelContextService = excelContextService ?? throw new ArgumentNullException(nameof(excelContextService));
            this.excelCommandExecutor = excelCommandExecutor ?? throw new ArgumentNullException(nameof(excelCommandExecutor));
            this.agentOrchestrator = agentOrchestrator ?? throw new ArgumentNullException(nameof(agentOrchestrator));
        }

        public string Route(string rawRequestJson)
        {
            var response = RouteInternal(rawRequestJson);
            return JsonConvert.SerializeObject(response, SerializerSettings);
        }

        public async Task<string> RouteAsync(string rawRequestJson)
        {
            WebMessageRequest request;
            try
            {
                request = JsonConvert.DeserializeObject<WebMessageRequest>(rawRequestJson);
            }
            catch (JsonException)
            {
                return Route(rawRequestJson);
            }

            if (request == null ||
                string.IsNullOrWhiteSpace(request.Type) ||
                string.IsNullOrWhiteSpace(request.RequestId))
            {
                return Route(rawRequestJson);
            }

            if (!string.Equals(request.Type, BridgeMessageTypes.RunAgent, StringComparison.Ordinal) &&
                !string.Equals(request.Type, BridgeMessageTypes.RunSkill, StringComparison.Ordinal))
            {
                return Route(rawRequestJson);
            }

            OfficeAgentLog.Info("bridge", "request.received", $"Received {request.Type}.", request.RequestId);

            WebMessageResponse response;
            try
            {
                response = string.Equals(request.Type, BridgeMessageTypes.RunSkill, StringComparison.Ordinal)
                    ? await RunSkillAsync(request).ConfigureAwait(true)
                    : await RunAgentAsync(request).ConfigureAwait(true);
            }
            catch (Exception error)
            {
                OfficeAgentLog.Error("bridge", "request.failed", $"Unhandled bridge failure while processing {request.Type}.", error, request.RequestId);
                response = Error(
                    request.Type,
                    request.RequestId,
                    code: "internal_error",
                    message: "OfficeAgent hit an unexpected error. Check the local log and try again.");
            }

            return JsonConvert.SerializeObject(response, SerializerSettings);
        }

        private WebMessageResponse RouteInternal(string rawRequestJson)
        {
            WebMessageRequest request;

            try
            {
                request = JsonConvert.DeserializeObject<WebMessageRequest>(rawRequestJson);
            }
            catch (JsonException)
            {
                return Error(
                    type: "bridge.unknown",
                    requestId: string.Empty,
                    code: "malformed_json",
                    message: "The web message payload was not valid JSON.");
            }

            if (request == null ||
                string.IsNullOrWhiteSpace(request.Type) ||
                string.IsNullOrWhiteSpace(request.RequestId))
            {
                return Error(
                    type: request?.Type ?? "bridge.unknown",
                    requestId: request?.RequestId ?? string.Empty,
                    code: "malformed_request",
                    message: "Web messages must include both type and requestId.");
            }

            if (!allowedTypes.Contains(request.Type))
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "unknown_message",
                    message: $"Message type '{request.Type}' is not allowed.");
            }

            OfficeAgentLog.Info("bridge", "request.received", $"Received {request.Type}.", request.RequestId);

            try
            {
                switch (request.Type)
                {
                    case BridgeMessageTypes.Ping:
                        if (HasUnexpectedPayload(request.Payload))
                        {
                            return Error(
                                request.Type,
                                request.RequestId,
                                code: "malformed_payload",
                                message: "bridge.ping does not accept a payload.");
                        }

                        return Success(
                            request.Type,
                            request.RequestId,
                            new PingPayload
                            {
                                Host = "OfficeAgent.ExcelAddIn",
                                Version = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "dev",
                            });
                    case BridgeMessageTypes.GetSettings:
                        if (HasUnexpectedPayload(request.Payload))
                        {
                            return Error(
                                request.Type,
                                request.RequestId,
                                code: "malformed_payload",
                                message: "bridge.getSettings does not accept a payload.");
                        }

                        return Success(request.Type, request.RequestId, settingsStore.Load());
                    case BridgeMessageTypes.GetSessions:
                        if (HasUnexpectedPayload(request.Payload))
                        {
                            return Error(
                                request.Type,
                                request.RequestId,
                                code: "malformed_payload",
                                message: "bridge.getSessions does not accept a payload.");
                        }

                        return Success(request.Type, request.RequestId, sessionStore.Load());
                    case BridgeMessageTypes.GetSelectionContext:
                        if (HasUnexpectedPayload(request.Payload))
                        {
                            return Error(
                                request.Type,
                                request.RequestId,
                                code: "malformed_payload",
                                message: "bridge.getSelectionContext does not accept a payload.");
                        }

                        return Success(request.Type, request.RequestId, excelContextService.GetCurrentSelectionContext());
                    case BridgeMessageTypes.SaveSettings:
                        return SaveSettings(request);
                    case BridgeMessageTypes.ExecuteExcelCommand:
                        return ExecuteExcelCommand(request);
                    case BridgeMessageTypes.RunSkill:
                        return RunSkill(request);
                    case BridgeMessageTypes.RunAgent:
                        return RunAgent(request);
                    default:
                        return Error(
                            request.Type,
                            request.RequestId,
                            code: "unknown_message",
                            message: $"Message type '{request.Type}' is not allowed.");
                }
            }
            catch (Exception error)
            {
                OfficeAgentLog.Error("bridge", "request.failed", $"Unhandled bridge failure while processing {request.Type}.", error, request.RequestId);
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "internal_error",
                    message: "OfficeAgent hit an unexpected error. Check the local log and try again.");
            }
        }

        private WebMessageResponse RunSkill(WebMessageRequest request)
        {
            if (request.Payload == null || request.Payload.Type != JTokenType.Object || !request.Payload.HasValues)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.runSkill requires a skill payload.");
            }

            try
            {
                var envelope = request.Payload.ToObject<AgentCommandEnvelope>() ?? new AgentCommandEnvelope();
                envelope.DispatchMode = AgentDispatchModes.Skill;
                return Success(request.Type, request.RequestId, agentOrchestrator.Execute(envelope));
            }
            catch (JsonException)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.runSkill requires a valid skill payload.");
            }
            catch (ArgumentException error)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "invalid_command",
                    message: error.Message);
            }
            catch (InvalidOperationException error)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "skill_failed",
                    message: error.Message);
            }
        }

        private async Task<WebMessageResponse> RunSkillAsync(WebMessageRequest request)
        {
            if (request.Payload == null || request.Payload.Type != JTokenType.Object || !request.Payload.HasValues)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.runSkill requires a skill payload.");
            }

            try
            {
                var envelope = request.Payload.ToObject<AgentCommandEnvelope>() ?? new AgentCommandEnvelope();
                envelope.DispatchMode = AgentDispatchModes.Skill;
                var result = await agentOrchestrator.ExecuteAsync(envelope).ConfigureAwait(true);
                return Success(request.Type, request.RequestId, result);
            }
            catch (JsonException)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.runSkill requires a valid skill payload.");
            }
            catch (ArgumentException error)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "invalid_command",
                    message: error.Message);
            }
            catch (InvalidOperationException error)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "skill_failed",
                    message: error.Message);
            }
        }

        private WebMessageResponse RunAgent(WebMessageRequest request)
        {
            if (request.Payload == null || request.Payload.Type != JTokenType.Object || !request.Payload.HasValues)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.runAgent requires an agent payload.");
            }

            try
            {
                var envelope = request.Payload.ToObject<AgentCommandEnvelope>() ?? new AgentCommandEnvelope();
                envelope.DispatchMode = AgentDispatchModes.Agent;
                return Success(request.Type, request.RequestId, agentOrchestrator.Execute(envelope));
            }
            catch (JsonException)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.runAgent requires a valid agent payload.");
            }
            catch (ArgumentException error)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "invalid_command",
                    message: error.Message);
            }
            catch (InvalidOperationException error)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "agent_failed",
                    message: error.Message);
            }
        }

        private async Task<WebMessageResponse> RunAgentAsync(WebMessageRequest request)
        {
            if (request.Payload == null || request.Payload.Type != JTokenType.Object || !request.Payload.HasValues)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.runAgent requires an agent payload.");
            }

            try
            {
                var envelope = request.Payload.ToObject<AgentCommandEnvelope>() ?? new AgentCommandEnvelope();
                envelope.DispatchMode = AgentDispatchModes.Agent;
                var result = await agentOrchestrator.ExecuteAsync(envelope).ConfigureAwait(true);
                return Success(request.Type, request.RequestId, result);
            }
            catch (JsonException)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.runAgent requires a valid agent payload.");
            }
            catch (ArgumentException error)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "invalid_command",
                    message: error.Message);
            }
            catch (InvalidOperationException error)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "agent_failed",
                    message: error.Message);
            }
        }

        private WebMessageResponse SaveSettings(WebMessageRequest request)
        {
            if (!HasValidSettingsPayload(request.Payload))
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.saveSettings requires a settings payload.");
            }

            try
            {
                var settings = request.Payload.ToObject<AppSettings>() ?? new AppSettings();
                settingsStore.Save(settings);
                return Success(request.Type, request.RequestId, settingsStore.Load());
            }
            catch (JsonException)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.saveSettings requires a valid settings payload.");
            }
        }

        private static bool HasValidSettingsPayload(JToken payload)
        {
            if (payload == null || payload.Type != JTokenType.Object || !payload.HasValues)
            {
                return false;
            }

            var payloadObject = (JObject)payload;
            return IsStringToken(payloadObject["apiKey"]) &&
                   IsStringToken(payloadObject["baseUrl"]) &&
                   IsStringToken(payloadObject["model"]) &&
                   payloadObject.Count >= 3;
        }

        private static bool HasUnexpectedPayload(JToken payload)
        {
            if (payload == null || payload.Type == JTokenType.Null)
            {
                return false;
            }

            return true;
        }

        private static bool IsStringToken(JToken token)
        {
            return token != null && token.Type == JTokenType.String;
        }

        private WebMessageResponse ExecuteExcelCommand(WebMessageRequest request)
        {
            if (request.Payload == null || request.Payload.Type != JTokenType.Object || !request.Payload.HasValues)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.executeExcelCommand requires a command payload.");
            }

            try
            {
                var command = request.Payload.ToObject<ExcelCommand>() ?? new ExcelCommand();
                confirmationService.Validate(command);

                var result = confirmationService.RequiresConfirmation(command) && !command.Confirmed
                    ? excelCommandExecutor.Preview(command)
                    : excelCommandExecutor.Execute(command);

                return Success(request.Type, request.RequestId, result);
            }
            catch (JsonException)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.executeExcelCommand requires a valid command payload.");
            }
            catch (ArgumentException error)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "invalid_command",
                    message: error.Message);
            }
            catch (InvalidOperationException error)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "command_failed",
                    message: error.Message);
            }
        }

        private static WebMessageResponse Success(string type, string requestId, object payload)
        {
            return new WebMessageResponse
            {
                Type = type,
                RequestId = requestId,
                Ok = true,
                Payload = payload,
            };
        }

        private static WebMessageResponse Error(string type, string requestId, string code, string message)
        {
            return new WebMessageResponse
            {
                Type = type,
                RequestId = requestId,
                Ok = false,
                Error = new WebMessageError
                {
                    Code = code,
                    Message = message,
                },
            };
        }
    }
}
