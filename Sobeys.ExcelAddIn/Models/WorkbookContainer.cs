using Sobeys.ExcelAddIn.Services;
using System.ComponentModel.Composition.Hosting;

namespace Sobeys.ExcelAddIn.Models
{
    public class WorkbookContainer
    {
        public WorkbookContainer(CompositionContainer container, WorkbookService workbookWrapper)
        {
            Container = container;
            WorkbookWrapper = workbookWrapper;
        }

        public CompositionContainer Container { get; }

        public WorkbookService WorkbookWrapper { get; }
    }
}
