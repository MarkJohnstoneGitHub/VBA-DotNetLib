// https://learn.microsoft.com/en-us/dotnet/api/system.io.streamwriter?view=netframework-4.8.1

using GIO = global::System.IO;
using GText = global::System.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Description("Implements a TextWriter for writing characters to a stream in a particular encoding.")]
    [Guid("A4DABE6B-485E-4AB0-8493-08E42A19343C")]
    [ProgId("DotNetLib.System.IO.StreamWriterSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IStreamWriterSingleton))]
    public class StreamWriterSingleton : IStreamWriterSingleton
    {
        public StreamWriterSingleton() { }

        // Factory Methods
        public StreamWriter Create(string path)
        {
            return new StreamWriter(path);
        }

        public StreamWriter Create(string path, bool append)
        {
            return new StreamWriter(path, append);
        }

        public StreamWriter Create(string path, bool append, GText.Encoding encoding)
        {
            return new StreamWriter(path, append, encoding);
        }

        public StreamWriter Create(string path, bool append, GText.Encoding encoding, int bufferSize)
        {
            return new StreamWriter(path, append, encoding, bufferSize);
        }

        public StreamWriter Create(GIO.Stream stream)
        {
            return new StreamWriter(stream);
        }

        public StreamWriter Create(GIO.Stream stream, GText.Encoding encoding)
        {
            return new StreamWriter(stream, encoding);
        }

        public StreamWriter Create(GIO.Stream stream, GText.Encoding encoding, int bufferSize)
        {
            return new StreamWriter(stream, encoding, bufferSize);
        }

        public StreamWriter Create(GIO.Stream stream, GText.Encoding encoding, int bufferSize, bool leaveOpen)
        {
            return new StreamWriter(stream, encoding, bufferSize, leaveOpen);
        }

        // Fields
        public StreamWriter NullStreamWriter => StreamWriter.Null;




    }
}
