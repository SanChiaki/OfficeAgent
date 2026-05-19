using System.Threading;
using System.Threading.Tasks;

namespace OfficeAgent.ExcelAddIn.Updates
{
    internal interface IUpdateManifestClient
    {
        Task<UpdateManifest> GetManifestAsync(string manifestUrl, CancellationToken cancellationToken);
    }
}
