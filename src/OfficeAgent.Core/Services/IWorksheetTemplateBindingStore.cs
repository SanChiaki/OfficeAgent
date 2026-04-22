using OfficeAgent.Core.Models;

namespace OfficeAgent.Core.Services
{
    public interface IWorksheetTemplateBindingStore
    {
        void SaveTemplateBinding(SheetTemplateBinding binding);

        SheetTemplateBinding LoadTemplateBinding(string sheetName);

        void ClearTemplateBinding(string sheetName);
    }
}
