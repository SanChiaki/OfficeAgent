using System.Threading.Tasks;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IAiColumnMappingClient
    {
        AiColumnMappingResponse Map(AiColumnMappingRequest request);

        Task<AiColumnMappingResponse> MapAsync(AiColumnMappingRequest request);
    }
}
