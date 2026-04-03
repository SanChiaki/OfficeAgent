using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IAgentFetchClient
    {
        Task<FetchResult> FetchAsync(string url, JObject headers = null);
    }
}
