using System.Collections.Generic;
using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface ITemplateCatalog
    {
        IReadOnlyList<TemplateDefinition> ListTemplates(string sheetName);

        SheetTemplateState GetSheetState(string sheetName);

        void ApplyTemplateToSheet(string sheetName, string templateId);

        void SaveSheetToExistingTemplate(string sheetName, string templateId, int expectedRevision, bool overwriteRevisionConflict);

        void SaveSheetAsNewTemplate(string sheetName, string templateName);

        void DetachTemplate(string sheetName);
    }
}
