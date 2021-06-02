using System;
using System.IO;

namespace Sobeys.ExcelAddIn.Updater
{
    public static class PathHelper
    {
        public static string GetInstallationPath()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Sobeys");
        }
    }
}
