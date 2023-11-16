// https://learn.microsoft.com/en-us/dotnet/api/system.io.fileinfo?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("7DCF6C2D-6060-41FB-AB74-C6AF40D417AA")]
    [Description("Provides properties and instance methods for the creation, copying, deletion, moving, and opening of files, and aids in the creation of FileStream objects. This class cannot be inherited.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IFileInfoSingleton
    {
        [Description("Initializes a new instance of the FileInfo class, which acts as a wrapper for a file path.")]
        FileInfo Create(string fileName);
    }
}
