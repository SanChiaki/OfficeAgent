using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IBusinessExportTemplateConnector
    {
        IReadOnlyList<BusinessExportTemplateOption> GetBusinessExportTemplates(string projectId);

        Task<BusinessExportWorkbook> ExportBusinessWorkbookAsync(
            string projectId,
            string templateId,
            CancellationToken cancellationToken);
    }
}
