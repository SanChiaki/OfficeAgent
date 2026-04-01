using System.Threading.Tasks;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IAgentFetchClient
    {
        Task<FetchResult> FetchAsync(string url);
    }
}
