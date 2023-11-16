using GIO = global::System.IO;
using DotNetLib.System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace DotNetLib.Extensions
{
    [ComVisible(false)]
    public static class DirectoryInfoExtension
    {
        public static List<DirectoryInfo> Wrap(IEnumerable<GIO.DirectoryInfo> directoryInfoList)
        {
            List<DirectoryInfo> wrappedDirectoryInfoList = new List<DirectoryInfo>(directoryInfoList.Count());
            foreach (GIO.DirectoryInfo directoryInfo in directoryInfoList)
            {
                wrappedDirectoryInfoList.Add(new DirectoryInfo(directoryInfo));
            }
            return wrappedDirectoryInfoList;
        }

    }
}
