using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;
using OfficeAgent.Core.Diagnostics;
using OfficeAgent.Core.Models;
using OfficeAgent.Core.Services;
using OfficeAgent.Infrastructure.Http;
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
            BridgeMessageTypes.SaveSessions,
            BridgeMessageTypes.SaveSettings,
            BridgeMessageTypes.ExecuteExcelCommand,
            BridgeMessageTypes.RunSkill,
            BridgeMessageTypes.RunAgent,
            BridgeMessageTypes.Login,
            BridgeMessageTypes.Logout,
            BridgeMessageTypes.GetLoginStatus,
        };
        private readonly FileSessionStore sessionStore;
        private readonly FileSettingsStore settingsStore;
        private readonly SharedCookieContainer sharedCookies;
        private readonly FileCookieStore cookieStore;

        public WebMessageRouter(
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            IExcelContextService excelContextService,
            IExcelCommandExecutor excelCommandExecutor,
            IAgentOrchestrator agentOrchestrator,
            SharedCookieContainer sharedCookies,
            FileCookieStore cookieStore)
        {
            this.sessionStore = sessionStore ?? throw new ArgumentNullException(nameof(sessionStore));
            this.settingsStore = settingsStore ?? throw new ArgumentNullException(nameof(settingsStore));
            this.excelContextService = excelContextService ?? throw new ArgumentNullException(nameof(excelContextService));
            this.excelCommandExecutor = excelCommandExecutor ?? throw new ArgumentNullException(nameof(excelCommandExecutor));
            this.agentOrchestrator = agentOrchestrator ?? throw new ArgumentNullException(nameof(agentOrchestrator));
            this.sharedCookies = sharedCookies ?? throw new ArgumentNullException(nameof(sharedCookies));
            this.cookieStore = cookieStore ?? throw new ArgumentNullException(nameof(cookieStore));
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
                !string.Equals(request.Type, BridgeMessageTypes.RunSkill, StringComparison.Ordinal) &&
                !string.Equals(request.Type, BridgeMessageTypes.Login, StringComparison.Ordinal))
            {
                return Route(rawRequestJson);
            }

            OfficeAgentLog.Info("bridge", "request.received", $"Received {request.Type}.", request.RequestId);

            WebMessageResponse response;
            try
            {
                if (string.Equals(request.Type, BridgeMessageTypes.RunSkill, StringComparison.Ordinal))
                {
                    response = await RunSkillAsync(request).ConfigureAwait(true);
                }
                else if (string.Equals(request.Type, BridgeMessageTypes.Login, StringComparison.Ordinal))
                {
                    response = await LoginAsync(request).ConfigureAwait(true);
                }
                else
                {
                    response = await RunAgentAsync(request).ConfigureAwait(true);
                }
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
                                Host = "Resy AI",
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
                    case BridgeMessageTypes.SaveSessions:
                        return SaveSessions(request);
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
                    case BridgeMessageTypes.GetLoginStatus:
                        return GetLoginStatus(request);
                    case BridgeMessageTypes.Logout:
                        return Logout(request);
                    case BridgeMessageTypes.Login:
                        return Error(
                            request.Type,
                            request.RequestId,
                            code: "invalid_dispatch",
                            message: "bridge.login must be routed asynchronously.");
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

        private WebMessageResponse GetLoginStatus(WebMessageRequest request)
        {
            if (HasUnexpectedPayload(request.Payload))
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.getLoginStatus does not accept a payload.");
            }

            var settings = settingsStore.Load();
            var ssoDomain = sharedCookies.SsoDomain;
            var isLoggedIn = false;

            if (!string.IsNullOrWhiteSpace(ssoDomain))
            {
                try
                {
                    var cookies = sharedCookies.Container.GetCookies(new Uri($"https://{ssoDomain}"));
                    isLoggedIn = cookies.Count > 0;
                }
                catch (UriFormatException)
                {
                    isLoggedIn = false;
                }
            }

            return Success(request.Type, request.RequestId, new LoginStatusPayload
            {
                IsLoggedIn = isLoggedIn,
                SsoUrl = settings.SsoUrl,
            });
        }

        private WebMessageResponse Logout(WebMessageRequest request)
        {
            if (HasUnexpectedPayload(request.Payload))
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.logout does not accept a payload.");
            }

            var ssoDomain = sharedCookies.SsoDomain;
            if (!string.IsNullOrWhiteSpace(ssoDomain))
            {
                try
                {
                    var cookies = sharedCookies.Container.GetCookies(new Uri($"https://{ssoDomain}"));
                    foreach (System.Net.Cookie cookie in cookies)
                    {
                        cookie.Expired = true;
                    }
                }
                catch (UriFormatException)
                {
                    // Ignore invalid domain.
                }
            }

            cookieStore.Clear();

            return Success(request.Type, request.RequestId, new LoginResultPayload { Success = true });
        }

#pragma warning disable CS1998 // ShowDialog is synchronous; async signature required for routing consistency.
        private async Task<WebMessageResponse> LoginAsync(WebMessageRequest request)
        {
            // Read SSO URL from the request payload first; fall back to persisted settings.
            var ssoUrl = string.Empty;
            var successPath = string.Empty;
            if (request.Payload != null && request.Payload.Type == JTokenType.Object)
            {
                try
                {
                    var loginPayload = request.Payload.ToObject<LoginPayload>();
                    if (loginPayload != null)
                    {
                        ssoUrl = loginPayload.SsoUrl ?? string.Empty;
                        successPath = loginPayload.SsoLoginSuccessPath ?? string.Empty;
                    }
                }
                catch (JsonException)
                {
                    // Ignore malformed payload; fall through to settings.
                }
            }

            if (string.IsNullOrWhiteSpace(ssoUrl))
            {
                var persisted = settingsStore.Load();
                ssoUrl = persisted.SsoUrl;
                successPath = persisted.SsoLoginSuccessPath;
            }

            if (string.IsNullOrWhiteSpace(ssoUrl))
            {
                return Error(request.Type, request.RequestId, "missing_sso_url", "\u8BF7\u5148\u914D\u7F6E SSO URL\u3002");
            }

            try
            {
                using (var popup = new SsoLoginPopup(ssoUrl, successPath, sharedCookies, cookieStore))
                {
                    await popup.InitializeAsync().ConfigureAwait(true);
                    var result = popup.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        return Success(request.Type, request.RequestId, new LoginResultPayload { Success = true });
                    }

                    return Success(request.Type, request.RequestId, new LoginResultPayload { Success = false, Error = "\u7528\u6237\u53D6\u6D88\u4E86\u767B\u5F55\u3002" });
                }
            }
            catch (Exception error)
            {
                return Error(request.Type, request.RequestId, "login_failed", error.Message);
            }
        }

        private WebMessageResponse SaveSessions(WebMessageRequest request)
        {
            if (request.Payload == null || request.Payload.Type != JTokenType.Object || !request.Payload.HasValues)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.saveSessions requires a session state payload.");
            }

            try
            {
                var state = request.Payload.ToObject<SessionState>() ?? new SessionState();
                sessionStore.Save(state);
                return Success(request.Type, request.RequestId, sessionStore.Load());
            }
            catch (JsonException)
            {
                return Error(
                    request.Type,
                    request.RequestId,
                    code: "malformed_payload",
                    message: "bridge.saveSessions requires a valid session state payload.");
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
