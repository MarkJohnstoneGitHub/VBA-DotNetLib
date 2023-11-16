// https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo?view=netframework-4.8.1

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Description("Exposes instance methods for creating, moving, and enumerating through directories and subdirectories. This class cannot be inherited.")]
    [Guid("3F95FF80-1A8F-4325-8B9A-3397B059B40C")]
    [ProgId("DotNetLib.System.IO.DirectoryInfoSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDirectoryInfoSingleton))]
    public class DirectoryInfoSingleton : IDirectoryInfoSingleton
    {
        public DirectoryInfo Create(string path)
        {
            return new DirectoryInfo(path);
        }

    }
}
