// https://learn.microsoft.com/en-us/dotnet/api/system.io.streamwriter?view=netframework-4.8.1

using GIO = global::System.IO;
using GText = global::System.Text;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("BE08EF64-CA41-4791-9F5D-BB12C106267E")]
    [Description("Implements a TextWriter for writing characters to a stream in a particular encoding.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStreamWriterSingleton
    {
        // Factory Methods

        [Description("Initializes a new instance of the StreamWriter class for the specified file by using the default encoding and buffer size.")]
        StreamWriter Create(string path);

        [Description("Initializes a new instance of the StreamWriter class for the specified file by using the default encoding and buffer size. If the file exists, it can be either overwritten or appended to. If the file does not exist, this constructor creates a new file.")]
        StreamWriter Create(string path, bool append);

        [Description("Initializes a new instance of the StreamWriter class for the specified file by using the specified encoding and default buffer size. If the file exists, it can be either overwritten or appended to. If the file does not exist, this constructor creates a new file.")]
        StreamWriter Create(string path, bool append, GText.Encoding encoding);

        [Description("Initializes a new instance of the StreamWriter class for the specified file on the specified path, using the specified encoding and buffer size. If the file exists, it can be either overwritten or appended to. If the file does not exist, this constructor creates a new file")]
        StreamWriter Create(string path, bool append, GText.Encoding encoding, int bufferSize);

        [Description("Initializes a new instance of the StreamWriter class for the specified stream by using UTF-8 encoding and the default buffer size.")]
        StreamWriter Create(GIO.Stream stream);

        [Description("Initializes a new instance of the StreamWriter class for the specified stream by using the specified encoding and the default buffer size.")]
        StreamWriter Create(GIO.Stream stream, GText.Encoding encoding);

        [Description("Initializes a new instance of the StreamWriter class for the specified stream by using the specified encoding and buffer size.")]
        StreamWriter Create(GIO.Stream stream, GText.Encoding encoding, int bufferSize);

        [Description("Initializes a new instance of the StreamWriter class for the specified stream by using the specified encoding and buffer size, and optionally leaves the stream open.")]
        StreamWriter Create(GIO.Stream stream, GText.Encoding encoding, int bufferSize, bool leaveOpen);

        // Fields
        StreamWriter NullStreamWriter
        {
            [Description("Provides a StreamWriter with no backing store that can be written to, but not read from.")]
            get;
        }

    }
}
