using System;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.IO;
using Sobeys.ExcelAddIn.Models;
using Sobeys.ExcelAddIn.Updater;
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
                case RibbonButtons.Update:
                    CheckForUpdate();
                    break;
                case RibbonButtons.About:
                    Process.Start("https://github.com/frederikstonge/sobeys-excel-addin");
                    break;
            }
        }

        public bool GetEnabled(Office.IRibbonControl control)
        {
            return control.Id switch
            {
                RibbonButtons.About => true,
                RibbonButtons.Update => true,
                _ => false
            };
        }

        private void CheckForUpdate()
        {
            try
            {
                var installationPath = PathHelper.GetInstallationPath();
                var version = typeof(AddIn).Assembly.GetName().Version;

                var folderPath = Path.Combine(
                    installationPath,
                    $"app-{version.ToString(3)}");

                var startInfo = new ProcessStartInfo();
                startInfo.UseShellExecute = true;
                startInfo.CreateNoWindow = true;
                startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                startInfo.WorkingDirectory = folderPath;
                startInfo.FileName = "Sobeys.ExcelAddIn.Updater.exe";
                Process proc = Process.Start(startInfo);
            }
            catch
            {
                // ignored
            }
        }
    }
}
