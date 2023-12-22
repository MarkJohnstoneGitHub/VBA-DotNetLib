// https://learn.microsoft.com/en-us/dotnet/api/system.io.bufferedstream?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Description("Adds a buffering layer to read and write operations on another stream. This class cannot be inherited.")]
    [Guid("B30CE6B2-1428-46DD-8F37-87872C011E3C")]
    [ProgId("DotNetLib.System.IO.BufferedStreamSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IBufferedStreamSingleton))]
    public class BufferedStreamSingleton : IBufferedStreamSingleton
    {
        public BufferedStreamSingleton() { }

        //public BufferedStream Create(Stream stream)
        //{
        //    return new BufferedStream(stream);
        //}

        public BufferedStream Create(Stream stream, int bufferSize = 4096)
        {
            return new BufferedStream(stream, bufferSize);
        }

    }
}
