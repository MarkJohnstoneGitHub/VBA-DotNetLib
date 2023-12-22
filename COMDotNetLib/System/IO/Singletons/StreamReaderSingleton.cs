// https://learn.microsoft.com/en-us/dotnet/api/system.io.streamreader?view=netframework-4.8.1

using Encoding = DotNetLib.System.Text.Encoding;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Description("Implements a TextReader that reads characters from a byte stream in a particular encoding.")]
    [Guid("E417E19A-FADE-4488-B798-8EB4D1B0AF94")]
    [ProgId("DotNetLib.System.IO.StreamReaderSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IStreamReaderSingleton))]
    public class StreamReaderSingleton : IStreamReaderSingleton
    {
        public StreamReaderSingleton() { }

        public StreamReader Create(string path)
        {
            return new StreamReader(path);
        }

        public StreamReader Create(string path, Encoding encoding)
        { 
            return new StreamReader(path, encoding); 
        }

        public StreamReader Create(string path, bool detectEncodingFromByteOrderMarks)
        {
            return new StreamReader(path, detectEncodingFromByteOrderMarks);
        }

        public StreamReader Create(string path, Encoding encoding, bool detectEncodingFromByteOrderMarks, int bufferSize)
        {
            return new StreamReader(path, encoding, detectEncodingFromByteOrderMarks, bufferSize);
        }
    }
}
