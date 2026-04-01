using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OfficeAgent.ExcelAddIn.WebBridge
{
    internal static class BridgeMessageTypes
    {
        public const string Ping = "bridge.ping";
        public const string GetSettings = "bridge.getSettings";
        public const string GetSelectionContext = "bridge.getSelectionContext";
        public const string SelectionContextChanged = "bridge.selectionContextChanged";
        public const string GetSessions = "bridge.getSessions";
        public const string SaveSessions = "bridge.saveSessions";
        public const string SaveSettings = "bridge.saveSettings";
        public const string ExecuteExcelCommand = "bridge.executeExcelCommand";
        public const string RunSkill = "bridge.runSkill";
        public const string RunAgent = "bridge.runAgent";
        public const string Login = "bridge.login";
        public const string Logout = "bridge.logout";
        public const string GetLoginStatus = "bridge.getLoginStatus";
    }

    internal sealed class WebMessageRequest
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("requestId")]
        public string RequestId { get; set; }

        [JsonProperty("payload")]
        public JToken Payload { get; set; }
    }

    internal sealed class WebMessageResponse
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("requestId")]
        public string RequestId { get; set; }

        [JsonProperty("ok")]
        public bool Ok { get; set; }

        [JsonProperty("payload")]
        public object Payload { get; set; }

        [JsonProperty("error")]
        public WebMessageError Error { get; set; }
    }

    internal sealed class WebMessageError
    {
        [JsonProperty("code")]
        public string Code { get; set; }

        [JsonProperty("message")]
        public string Message { get; set; }
    }

    internal sealed class PingPayload
    {
        [JsonProperty("host")]
        public string Host { get; set; }

        [JsonProperty("version")]
        public string Version { get; set; }
    }

    internal sealed class WebMessageEvent
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("payload")]
        public object Payload { get; set; }
    }

    internal sealed class LoginResultPayload
    {
        [JsonProperty("success")]
        public bool Success { get; set; }

        [JsonProperty("error")]
        public string Error { get; set; } = string.Empty;
    }

    internal sealed class LoginStatusPayload
    {
        [JsonProperty("isLoggedIn")]
        public bool IsLoggedIn { get; set; }

        [JsonProperty("ssoUrl")]
        public string SsoUrl { get; set; } = string.Empty;
    }

    internal sealed class LoginPayload
    {
        [JsonProperty("ssoUrl")]
        public string SsoUrl { get; set; } = string.Empty;
    }
}
