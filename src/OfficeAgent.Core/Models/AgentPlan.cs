using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;

namespace OfficeAgent.Core.Models
{
    public static class PlannerResponseModes
    {
        public const string Message = "message";
        public const string ReadStep = "read_step";
        public const string Plan = "plan";
    }

    public static class PlannerStepTypes
    {
        public const string ReadSelectionTable = ExcelCommandTypes.ReadSelectionTable;
        public const string ReadRange = ExcelCommandTypes.ReadRange;
        public const string FetchUrl = "fetch.url";
        public const string UploadData = "skill.upload_data";
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class PlannerRequest
    {
        public string SessionId { get; set; } = string.Empty;

        public string UserInput { get; set; } = string.Empty;

        public SelectionContext SelectionContext { get; set; }

        public PlannerObservation[] Observations { get; set; } = System.Array.Empty<PlannerObservation>();

        public ConversationTurn[] ConversationHistory { get; set; } = System.Array.Empty<ConversationTurn>();

        public string ApiBaseUrl { get; set; } = string.Empty;
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class PlannerObservation
    {
        public string Kind { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;

        public ExcelTableData Table { get; set; }

        public JToken Data { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class PlannerResponse
    {
        public string Mode { get; set; } = string.Empty;

        public string AssistantMessage { get; set; } = string.Empty;

        public PlannerStep Step { get; set; }

        public AgentPlan Plan { get; set; }
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class PlannerStep
    {
        public string Type { get; set; } = string.Empty;

        public JObject Args { get; set; } = new JObject();
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class AgentPlan
    {
        public string Summary { get; set; } = string.Empty;

        public AgentPlanStep[] Steps { get; set; } = System.Array.Empty<AgentPlanStep>();
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class AgentPlanStep
    {
        public string Type { get; set; } = string.Empty;

        public JObject Args { get; set; } = new JObject();
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class PlanExecutionJournal
    {
        public bool HasFailures { get; set; }

        public string ErrorMessage { get; set; } = string.Empty;

        public PlanExecutionJournalStep[] Steps { get; set; } = System.Array.Empty<PlanExecutionJournalStep>();
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class PlanExecutionJournalStep
    {
        public string Type { get; set; } = string.Empty;

        public string Title { get; set; } = string.Empty;

        public string Status { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;

        public string ErrorMessage { get; set; } = string.Empty;
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class ConversationTurn
    {
        public string Role { get; set; } = string.Empty;

        public string Content { get; set; } = string.Empty;
    }
}
