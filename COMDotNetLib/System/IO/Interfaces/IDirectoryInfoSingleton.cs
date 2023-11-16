// https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("9B9A5C7B-DE66-4C14-A8CD-6B2326A464E1")]
    [Description("Exposes instance methods for creating, moving, and enumerating through directories and subdirectories. This class cannot be inherited.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDirectoryInfoSingleton
    {
        [Description("Initializes a new instance of the DirectoryInfo class on the specified path.")]
        DirectoryInfo Create(string path);
    }
}
