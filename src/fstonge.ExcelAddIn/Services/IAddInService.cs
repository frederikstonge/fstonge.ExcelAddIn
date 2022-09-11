namespace fstonge.ExcelAddIn.Services
{
    public interface IAddInService
    {
        void OnAction(Microsoft.Office.Core.IRibbonControl control);

        bool GetEnabled(Microsoft.Office.Core.IRibbonControl control);
    }
}