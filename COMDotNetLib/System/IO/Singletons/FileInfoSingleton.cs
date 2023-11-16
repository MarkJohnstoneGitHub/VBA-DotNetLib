// https://learn.microsoft.com/en-us/dotnet/api/system.io.fileinfo?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Description("Provides properties and instance methods for the creation, copying, deletion, moving, and opening of files, and aids in the creation of FileStream objects. This class cannot be inherited.")]
    [Guid("267FB24F-2580-41D1-81CF-70EB5785DED7")]
    [ProgId("DotNetLib.System.IO.FileInfoSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IFileInfoSingleton))]

    public class FileInfoSingleton : IFileInfoSingleton
    {
        public FileInfo Create(string fileName)
        {
            return new FileInfo(fileName);
        }

    }
}
