using System.Windows.Forms;
using Microsoft.Office.Core;

namespace fstonge.ExcelAddIn.Services
{
    public interface ITaskPaneFactory
    {
        Microsoft.Office.Tools.CustomTaskPane CreateTaskPane(UserControl userControl, string title, object window, MsoCTPDockPosition position);
    }
}