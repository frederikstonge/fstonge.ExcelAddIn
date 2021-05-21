using System.ComponentModel.Composition;
using Sobeys.ExcelAddIn.Models;
using Office = Microsoft.Office.Core;

namespace Sobeys.ExcelAddIn.Services
{
    [Export(typeof(IAddInService))]
    [PartCreationPolicy(CreationPolicy.Shared)]
    public class AddInService : IAddInService
    {
        [ImportingConstructor]
        public AddInService()
        {
        }

        public void OnAction(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case RibbonButtons.About:
                    System.Diagnostics.Process.Start("https://github.com/frederikstonge/sobeys-excel-addin");
                    break;
            }
        }

        public bool GetEnabled(Office.IRibbonControl control)
        {
            return control.Id switch
            {
                RibbonButtons.About => true,
                _ => false
            };
        }
    }
}
