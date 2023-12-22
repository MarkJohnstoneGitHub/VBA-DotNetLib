// https://learn.microsoft.com/en-us/dotnet/api/system.io.streamreader?view=netframework-4.8.1

using Encoding = DotNetLib.System.Text.Encoding;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("52720B99-13E7-4D1D-8137-41E2DA73A184")]
    [Description("Implements a TextReader that reads characters from a byte stream in a particular encoding.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStreamReaderSingleton
    {
        [Description("Initializes a new instance of the StreamReader class for the specified file name.")]
        StreamReader Create(string path);

        [Description("Initializes a new instance of the StreamReader class for the specified file name, with the specified character encoding.")]
        StreamReader Create(string path, Encoding encoding);

        [Description("Initializes a new instance of the StreamReader class for the specified stream, with the specified byte order mark detection option.")]
        StreamReader Create(string path, bool detectEncodingFromByteOrderMarks);

        [Description("Initializes a new instance of the StreamReader class for the specified file name, with the specified character encoding and byte order mark detection option.")]
        StreamReader Create(string path, Encoding encoding, bool detectEncodingFromByteOrderMarks, int bufferSize);
    }
}
