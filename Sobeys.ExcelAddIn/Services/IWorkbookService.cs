namespace Sobeys.ExcelAddIn.Services
{
    public interface IWorkbookService
    {
        void OnAction(Microsoft.Office.Core.IRibbonControl control);

        bool GetEnabled(Microsoft.Office.Core.IRibbonControl control);

        void Dispose();
    }
}