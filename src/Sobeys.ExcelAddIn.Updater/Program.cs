using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using Squirrel;

namespace Sobeys.ExcelAddIn.Updater
{
    public class Program
    {
        private const string GithubUsername = "frederikstonge";
        private const string GithubProject = "sobeys-excel-addin";
        private const string RegistrySubKey = @"Software\Microsoft\Office\Excel\AddIns\Sobeys.ExcelAddIn";

        public static async Task Main()
        {
            try
            {
                using var mgr = await UpdateManager.GitHubUpdateManager($"https://github.com/{GithubUsername}/{GithubProject}");
                SquirrelAwareApp.HandleEvents(
                    onInitialInstall: v =>
                    {
                        mgr.CreateUninstallerRegistryEntry();
                        CreateRegistryEntries(v);
                    },
                    onAppUpdate: v =>
                    {
                        MessageBox.Show("Excel will close to update Sobeys Excel AddIn.", "Update");
                        StopExcel();
                        RemoveRegistryEntries();
                        CreateRegistryEntries(v);
                        StartExcel();
                    },
                    onAppUninstall: v =>
                    {
                        StopExcel();
                        RemoveRegistryEntries();
                        mgr.RemoveUninstallerRegistryEntry();
                    },
                    onFirstRun: StartExcel);

                await mgr.UpdateApp();
            }
            catch
            {
                // ignored
            }
        }

        private static void StartExcel()
        {
            Process.Start("excel.exe");
        }

        private static void StopExcel()
        {
            var processes = Process.GetProcessesByName("EXCEL");
            foreach (var process in processes)
            {
                process.Kill();
            }
        }

        private static void CreateRegistryEntries(Version version)
        {
            var manifestPath = Path.Combine(
                PathHelper.GetInstallationPath(),
                $"app-{version.ToString(3)}",
                "Sobeys.ExcelAddIn.vsto")
                + "|vstolocal";

            using var subKey = Registry.CurrentUser.CreateSubKey(RegistrySubKey, true);
            subKey.SetValue("FriendlyName", "Sobeys Excel Add-In");
            subKey.SetValue("Description", "Sobeys Excel Add-In");
            subKey.SetValue("Manifest", manifestPath);
            subKey.SetValue("LoadBehavior", 3);
        }

        private static void RemoveRegistryEntries()
        {
            Registry.CurrentUser.DeleteSubKeyTree(RegistrySubKey, false);
        }
    }
}