using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Octokit;

namespace Sobeys.ExcelAddIn
{
    public partial class ThisAddIn
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

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                var client = new GitHubClient(new ProductHeaderValue(GithubProject));
                var releases = client.Repository.Release.GetAll(GithubUsername, GithubProject).Result;

                var latestRelease = releases[0];

                var latestGitHubVersion = Version.Parse(latestRelease.TagName);
                var localVersion = Assembly.GetAssembly(typeof(ThisAddIn)).GetName().Version;

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

            _bootstrapper = new Bootstrapper(_ribbon);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _bootstrapper.Dispose();
        }

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
    }
}
