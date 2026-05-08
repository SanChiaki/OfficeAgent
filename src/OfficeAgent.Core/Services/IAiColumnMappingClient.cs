using System.Threading.Tasks;
using System.Threading;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IAiColumnMappingClient
    {
        AiColumnMappingResponse Map(AiColumnMappingRequest request);

        Task<AiColumnMappingResponse> MapAsync(AiColumnMappingRequest request);

        Task<AiColumnMappingResponse> MapAsync(AiColumnMappingRequest request, CancellationToken cancellationToken);
    }
}
