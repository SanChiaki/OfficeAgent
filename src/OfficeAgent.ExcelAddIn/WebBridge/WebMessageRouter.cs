using System;
using System.Collections.Generic;
using System.Reflection;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;
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
        private readonly HashSet<string> allowedTypes = new HashSet<string>(StringComparer.Ordinal)
        {
            BridgeMessageTypes.Ping,
            BridgeMessageTypes.GetSettings,
            BridgeMessageTypes.GetSelectionContext,
            BridgeMessageTypes.GetSessions,
            BridgeMessageTypes.SaveSettings,
            BridgeMessageTypes.ExecuteExcelCommand,
            BridgeMessageTypes.RunSkill,
        };
        private readonly FileSessionStore sessionStore;
        private readonly FileSettingsStore settingsStore;

        public WebMessageRouter(
            FileSessionStore sessionStore,
            FileSettingsStore settingsStore,
            IExcelContextService excelContextService)
        {
            this.sessionStore = sessionStore ?? throw new ArgumentNullException(nameof(sessionStore));
            this.settingsStore = settingsStore ?? throw new ArgumentNullException(nameof(settingsStore));
            this.excelContextService = excelContextService ?? throw new ArgumentNullException(nameof(excelContextService));
        }

        public string Route(string rawRequestJson)
        {
            var response = RouteInternal(rawRequestJson);
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
                case BridgeMessageTypes.RunSkill:
                    return Error(
                        request.Type,
                        request.RequestId,
                        code: "not_implemented",
                        message: $"Message type '{request.Type}' is registered but not implemented yet.");
                default:
                    return Error(
                        request.Type,
                        request.RequestId,
                        code: "unknown_message",
                        message: $"Message type '{request.Type}' is not allowed.");
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
