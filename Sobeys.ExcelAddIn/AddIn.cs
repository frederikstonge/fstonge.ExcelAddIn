using System;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Octokit;

namespace Sobeys.ExcelAddIn
{
    public partial class AddIn
    {
        private const string GithubUsername = "frederikstonge";
        private const string GithubProject = "sobeys-excel-addin";

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
            try
            {
                var client = new GitHubClient(new ProductHeaderValue(GithubProject));
                var releases = client.Repository.Release.GetAll(GithubUsername, GithubProject).Result;

                var latestRelease = releases[0];

                var latestGitHubVersion = Version.Parse(latestRelease.TagName);
                var localVersion = Assembly.GetAssembly(typeof(AddIn)).GetName().Version;

                int versionComparison = localVersion.CompareTo(latestGitHubVersion);
                if (versionComparison < 0)
                {
                    var result = MessageBox.Show(
                        Properties.Resources.NewVersionMessage,
                        Properties.Resources.NewVersionTitle,
                        MessageBoxButtons.YesNo);

                    if (result == DialogResult.Yes)
                    {
                        Process.Start($"https://github.com/{GithubUsername}/{GithubProject}/releases/tag/{latestRelease.TagName}");
                    }
                }
            }
            catch
            {
                // ignored
            }
        }

        private void InternalStartup()
        {
            Startup += AddIn_Startup;
            Shutdown += AddIn_Shutdown;
        }
    }
}
