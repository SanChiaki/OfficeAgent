using System.Threading.Tasks;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface ILlmPlannerClient
    {
        string Complete(PlannerRequest request);
        Task<string> CompleteAsync(PlannerRequest request);
    }
}
