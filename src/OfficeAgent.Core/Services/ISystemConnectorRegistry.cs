using System.Collections.Generic;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface ISystemConnectorRegistry
    {
        IReadOnlyList<ProjectOption> GetProjects();

        ISystemConnector GetRequiredConnector(string systemKey);
    }
}
