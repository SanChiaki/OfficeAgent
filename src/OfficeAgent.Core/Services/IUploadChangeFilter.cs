using System.Collections.Generic;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IUploadChangeFilter
    {
        UploadChangeFilterResult FilterUploadChanges(string projectId, IReadOnlyList<CellChange> changes);
    }
}
