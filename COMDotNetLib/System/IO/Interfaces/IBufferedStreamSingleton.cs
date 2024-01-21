// https://learn.microsoft.com/en-us/dotnet/api/system.io.bufferedstream?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("46CC7766-EC13-4FF3-ACCE-526AD703F6CA")]
    [Description("Adds a buffering layer to read and write operations on another stream. This class cannot be inherited.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IBufferedStreamSingleton
    {
        //[Description("Initializes a new instance of the BufferedStream class with a default buffer size of 4096 bytes.")]
        //BufferedStream Create2(Stream stream);

        [Description("Initializes a new instance of the BufferedStream class with the specified buffer size or default buffer size of 4096 byte.")]
        BufferedStream Create(Stream stream, int bufferSize = 4096);

    }
}
