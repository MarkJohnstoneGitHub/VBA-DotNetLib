using GIO = global::System.IO;
using DotNetLib.System.IO;
using System.Runtime.InteropServices;

namespace DotNetLib.Extensions
{
    [ComVisible(false)]
    public static class FileSystemInfoExtension
    {
       public static FileSystemInfo Wrap(this GIO.FileSystemInfo fileSystemInfo)
       {
            if (fileSystemInfo == null) { return null; }

            if (fileSystemInfo is GIO.DirectoryInfo directoryInfo)
            {
                return new DirectoryInfo(directoryInfo);
            }
            if (fileSystemInfo is GIO.FileInfo fileInfo)
            {
                return new FileInfo(fileInfo);
            }
            return null; //If not implemented return null
        }

    }
}
