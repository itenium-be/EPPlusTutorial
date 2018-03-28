using System;
using System.Diagnostics;
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
            var mth = new StackTrace().GetFrame(1).GetMethod();
            var cls = mth.ReflectedType.Name;

            var dir = new DirectoryInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excels"));
            Directory.CreateDirectory(dir.FullName);

            var name = cls + (fileName ?? callerName);
            if (!name.Contains("."))
            {
                name += ".xlsx";
            }
            return Path.Combine(dir.FullName, name);
        }
    }
}