using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;
using Squirrel;

namespace fstonge.ExcelAddIn.Updater
{
    public class Program
    {
        public const string GithubUsername = "frederikstonge";
        public const string GithubProject = "fstonge.ExcelAddIn";
        private const string RegistrySubKey = @"Software\Microsoft\Office\Excel\AddIns\fstonge.ExcelAddIn";

        public static async Task Main()
        {
            try
            {
                using var mgr = await UpdateManager.GitHubUpdateManager($"https://github.com/{GithubUsername}/{GithubProject}");
                SquirrelAwareApp.HandleEvents(
                    onInitialInstall: v =>
                    {
                        StopExcel();
                        mgr.CreateUninstallerRegistryEntry();
                        CreateRegistryEntries(v);
                        SetVstoDebugEnvironmentVariables();
                    },
                    onAppUpdate: v =>
                    {
                        MessageBox.Show("Excel will close to update fstonge Excel AddIn.", "Update");
                        StopExcel();
                        RemoveRegistryEntries();
                        CreateRegistryEntries(v);
                        StartExcel();
                    },
                    onAppUninstall: v =>
                    {
                        StopExcel();
                        RemoveVstoDebugEnvironmentVariables();
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
                "fstonge.ExcelAddIn.vsto")
                + "|vstolocal";

            using var subKey = Registry.CurrentUser.CreateSubKey(RegistrySubKey, true);
            subKey.SetValue("FriendlyName", "fstonge Excel Add-In");
            subKey.SetValue("Description", "fstonge Excel Add-In");
            subKey.SetValue("Manifest", manifestPath);
            subKey.SetValue("LoadBehavior", 3);
        }

        private static void RemoveRegistryEntries()
        {
            Registry.CurrentUser.DeleteSubKeyTree(RegistrySubKey, false);
        }

        private static void SetVstoDebugEnvironmentVariables()
        {
            Environment.SetEnvironmentVariable("VSTO_LOGALERTS", "1", EnvironmentVariableTarget.User);
            Environment.SetEnvironmentVariable("VSTO_SUPPRESSDISPLAYALERTS", "0", EnvironmentVariableTarget.User);
        }

        private static void RemoveVstoDebugEnvironmentVariables()
        {
            Environment.SetEnvironmentVariable("VSTO_LOGALERTS", null, EnvironmentVariableTarget.User);
            Environment.SetEnvironmentVariable("VSTO_SUPPRESSDISPLAYALERTS", null, EnvironmentVariableTarget.User);
        }
    }
}