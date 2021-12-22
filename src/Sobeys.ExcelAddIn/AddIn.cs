using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Microsoft.Office.Core;
using Sobeys.ExcelAddIn.Updater;

namespace Sobeys.ExcelAddIn
{
    public partial class AddIn
    {
        private Ribbon _ribbon;
        private Bootstrapper _bootstrapper;

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new Ribbon();
            return _ribbon;
        }

        private void AddIn_Startup(object sender, EventArgs e)
        {
            SetupLanguage();
            ValidateNewerVersion();
            _bootstrapper = new Bootstrapper(_ribbon);
        }

        private void AddIn_Shutdown(object sender, EventArgs e)
        {
            _bootstrapper.Dispose();
        }

        private void SetupLanguage()
        {
            var lcid = Globals.AddIn.Application.LanguageSettings.LanguageID[MsoAppLanguageID.msoLanguageIDUI];
            var culture = new CultureInfo(lcid);
            System.Threading.Thread.CurrentThread.CurrentUICulture = culture;
            System.Threading.Thread.CurrentThread.CurrentCulture = culture;
        }

        private void ValidateNewerVersion()
        {
            var shouldUpdate = Application.Visible;
#if DEBUG
            shouldUpdate = false;
#endif

            if (shouldUpdate)
            {
                try
                {
                    var installationPath = PathHelper.GetInstallationPath();

                    var version = typeof(AddIn).Assembly.GetName().Version;

                    var path = Path.Combine(
                        installationPath, 
                        $"app-{version.ToString(3)}",
                        "Sobeys.ExcelAddIn.Updater.exe");

                    var process = new Process
                    {
                        StartInfo =
                        {
                            UseShellExecute = true,
                            CreateNoWindow = true,
                            WindowStyle = ProcessWindowStyle.Hidden,
                            FileName = path,
                            RedirectStandardError = false,
                            RedirectStandardOutput = false,
                            WorkingDirectory = installationPath,
                        }
                    };
                    process.Start();
                }
                catch (Exception)
                {
                    // ignored
                }
            }
        }

        private void InternalStartup()
        {
            Startup += AddIn_Startup;
            Shutdown += AddIn_Shutdown;
        }
    }
}
