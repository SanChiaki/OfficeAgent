using System.Threading.Tasks;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IAgentOrchestrator
    {
        AgentCommandResult Execute(AgentCommandEnvelope envelope);
        Task<AgentCommandResult> ExecuteAsync(AgentCommandEnvelope envelope);
    }
}
