namespace Sobeys.ExcelAddIn.Services
{
    public interface IWorkbookService
    {
        void OnAction(Microsoft.Office.Core.IRibbonControl control);

        void OnPressedAction(Microsoft.Office.Core.IRibbonControl control, bool isPressed);

        bool GetEnabled(Microsoft.Office.Core.IRibbonControl control);

        bool GetPressed(Microsoft.Office.Core.IRibbonControl control);

        void Dispose();
    }
}