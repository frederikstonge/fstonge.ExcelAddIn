using Microsoft.Office.Core;
using Octokit;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;

namespace Sobeys.ExcelAddIn
{
    public partial class ThisAddIn
    {
        private const string GithubUsername = "frederikstonge";
        private const string GithubProject = "sobeys-excel-addin";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
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

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
