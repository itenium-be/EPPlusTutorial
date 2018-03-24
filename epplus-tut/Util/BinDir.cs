using System;
using System.IO;
using System.Runtime.CompilerServices;

namespace EPPlusTutorial.Util
{
    public static class BinDir
    {
        /// <summary>
        /// Save Excels from the UnitTests under the bin/excels folder of this project
        /// </summary>
        public static string GetPath(string fileName = null, [CallerMemberName] string callerName = "")
        {
            // TODO: Prefix filename with calling className
            var dir = new DirectoryInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excels"));
            Directory.CreateDirectory(dir.FullName);

            var name = fileName ?? callerName;
            if (!name.Contains("."))
            {
                name += ".xlsx";
            }
            return Path.Combine(dir.FullName, name);
        }
    }
}