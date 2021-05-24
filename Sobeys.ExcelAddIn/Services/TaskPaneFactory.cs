using System.ComponentModel.Composition;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Tools = Microsoft.Office.Tools;

namespace Sobeys.ExcelAddIn.Services
{
    [Export(typeof(ITaskPaneFactory))]
    public class TaskPaneFactory : ITaskPaneFactory
    {
        private readonly ThisAddIn _thisAddIn;

        [ImportingConstructor]
        public TaskPaneFactory(ThisAddIn thisAddIn)
        {
            _thisAddIn = thisAddIn;
        }

        public Tools.CustomTaskPane CreateTaskPane(UserControl userControl, string title, object window, MsoCTPDockPosition position)
        {
            var taskPane = _thisAddIn.CustomTaskPanes.Add(userControl, title, window);
            taskPane.DockPosition = position;
            taskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            return taskPane;
        }
    }
}
