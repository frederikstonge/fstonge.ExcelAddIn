using Microsoft.Office.Core;
using Octokit;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;

namespace Sobeys.ExcelAddIn
{
    public partial class ThisAddIn
    {
        private const string GithubUsername = "frederikstonge";
        private const string GithubProject = "sobeys-excel-addin";

        private Ribbon _ribbon;
        public AddInWrapper AddInWrapper { get; private set; }


        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new Ribbon(this);
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
                    var result = MessageBox.Show("A newer version of the addin was detected, do you wish to download it?", "New version detected", MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                    {
                        Process.Start($"https://github.com/{GithubUsername}/{GithubProject}/releases/tag/{latestRelease.TagName}");
                    }
                }
            }
            catch
            {

            }


            AddInWrapper = new AddInWrapper(_ribbon);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            AddInWrapper.Dispose();
        }

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}
