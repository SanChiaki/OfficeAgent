using System.Collections.Generic;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface ISystemConnector
    {
        IReadOnlyList<ProjectOption> GetProjects();

        WorksheetSchema GetSchema(string projectId);

        IReadOnlyList<IDictionary<string, object>> Find(string projectId, IReadOnlyList<string> rowIds, IReadOnlyList<string> fieldKeys);

        void BatchSave(string projectId, IReadOnlyList<CellChange> changes);
    }
}
