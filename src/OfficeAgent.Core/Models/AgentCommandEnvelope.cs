using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace OfficeAgent.Core.Models
{
    public static class AgentRouteTypes
    {
        public const string Chat = "chat";
        public const string ExcelCommand = "excelCommand";
        public const string Skill = "skill";
        public const string Plan = "plan";
    }

    public static class AgentDispatchModes
    {
        public const string Auto = "auto";
        public const string Agent = "agent";
        public const string Skill = "skill";
    }

    public static class SkillNames
    {
        public const string UploadData = "upload_data";
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class AgentCommandEnvelope
    {
        public string DispatchMode { get; set; } = AgentDispatchModes.Auto;

        public string SessionId { get; set; } = string.Empty;

        public string UserInput { get; set; } = string.Empty;

        public string SkillName { get; set; } = string.Empty;

        public bool Confirmed { get; set; }

        public AgentPlan Plan { get; set; }

        public UploadPreview UploadPreview { get; set; }

        public ConversationTurn[] ConversationHistory { get; set; } = System.Array.Empty<ConversationTurn>();
    }

    [JsonObject(NamingStrategyType = typeof(CamelCaseNamingStrategy))]
    public sealed class AgentCommandResult
    {
        public string Route { get; set; } = string.Empty;

        public string SkillName { get; set; } = string.Empty;

        public bool RequiresConfirmation { get; set; }

        public string Status { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;

        public PlannerResponse Planner { get; set; }

        public PlanExecutionJournal Journal { get; set; }

        public ExcelCommandPreview Preview { get; set; }

        public UploadPreview UploadPreview { get; set; }
    }
}
