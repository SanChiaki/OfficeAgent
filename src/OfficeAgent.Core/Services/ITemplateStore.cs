using System.Collections.Generic;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface ITemplateStore
    {
        IReadOnlyList<TemplateDefinition> ListByProject(string systemKey, string projectId);

        TemplateDefinition Load(string templateId);

        TemplateDefinition SaveNew(TemplateDefinition template);

        TemplateDefinition SaveExisting(TemplateDefinition template, int expectedRevision);
    }
}
