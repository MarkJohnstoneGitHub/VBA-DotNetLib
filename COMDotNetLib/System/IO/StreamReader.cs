// https://learn.microsoft.com/en-us/dotnet/api/system.io.streamreader?view=netframework-4.8.1

using GSystem = global::System;
using GIO = global::System.IO;
using GText = global::System.Text;

using System;
using System.Runtime.InteropServices;
using System.Text;
using DotNetLib.Extensions;
using System.ComponentModel;
using DotNetLib.System.Text;
using Encoding = DotNetLib.System.Text.Encoding;
using System.IO;
using System.Runtime.InteropServices.ComTypes;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Description("Implements a TextReader that reads characters from a byte stream in a particular encoding.")]
    [Guid("78B5CEA2-53E4-4CBD-A025-40DFEED4E8C2")]
    [ProgId("DotNetLib.System.IO.StreamReader")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IStreamReader))]
    public class StreamReader : GIO.TextReader, IStreamReader, TextReader, IDisposable, IWrappedObject
    {
        private GIO.StreamReader _streamReader;

        private Stream _baseStream;

        private Encoding _encoding;

        // Constructors
        public StreamReader(string path)
        {
            _streamReader = new GIO.StreamReader(path);
        }

        public StreamReader(string path, GText.Encoding encoding)
        {
            _streamReader = new GIO.StreamReader(path, encoding);
        }

        public StreamReader(string path, Encoding encoding)
        {
            _streamReader = new GIO.StreamReader(path, encoding.UnWrapEncoding());
        }

        public StreamReader(string path, Encoding encoding, bool detectEncodingFromByteOrderMarks, int bufferSize)
        {
            _streamReader = new GIO.StreamReader(path, encoding.UnWrapEncoding(), detectEncodingFromByteOrderMarks, bufferSize);
        }
        public StreamReader(string path, bool detectEncodingFromByteOrderMarks)
        {
            _streamReader = new GIO.StreamReader(path, detectEncodingFromByteOrderMarks);
        }

        public StreamReader(string path, GText.Encoding encoding, bool detectEncodingFromByteOrderMarks)
        {
            _streamReader = new GIO.StreamReader(path, encoding, detectEncodingFromByteOrderMarks);
        }


        public StreamReader(GIO.Stream stream)
        {
            _streamReader = new GIO.StreamReader(stream);
        }

        public StreamReader(Stream stream)
        {
            _streamReader = new GIO.StreamReader((GIO.Stream)stream);
        }


        // Properties

        public object WrappedObject => _streamReader;

        internal GIO.StreamReader WrappedStreamReader => _streamReader;

        public Stream BaseStream => throw new NotImplementedException();

        public Encoding CurrentEncoding
        {
            get
            {
                if (_encoding == null)
                    _encoding = _streamReader.CurrentEncoding.Wrap();
                return _encoding;
            }
        }

        public bool EndOfStream => _streamReader.EndOfStream;



        // Methods

        public new void Close()
        {
            _streamReader.Close();
        }

        public new void Dispose()
        {
            _streamReader.Dispose();
        }

        public new int Peek()
        {
            return _streamReader.Peek();
        }

        public new int Read()
        {
            return _streamReader.Read();
        }

        public int Read([In, Out] ref byte[] buffer, int index, int count)
        {
            // https://stackoverflow.com/questions/5431004/convert-byte-to-char
            char[] chars = _streamReader.CurrentEncoding.GetChars(buffer,index,count);  // Todo check implementation
            int numberOfChars = _streamReader.Read(chars, index, count);
            buffer = _streamReader.CurrentEncoding.GetBytes(chars);
            return numberOfChars;
        }

        public int ReadBlock([In, Out] ref byte[] buffer, int index, int count)
        {
            // https://stackoverflow.com/questions/5431004/convert-byte-to-char
            char[] chars = _streamReader.CurrentEncoding.GetChars(buffer, index, count); // Todo check implementation
            int numberOfChars = _streamReader.ReadBlock(chars, index, count);
            buffer = _streamReader.CurrentEncoding.GetBytes(chars);
            return numberOfChars;
        }

        public new string ReadLine()
        {
            return _streamReader.ReadLine();
        }

        public new string ReadToEnd()
        {
            return _streamReader.ReadToEnd();
        }

        public new Type GetType()
        {
            return new Type(((GSystem.Object)this).GetType());
        }

        public new virtual string ToString()
        { 
            return _streamReader.ToString(); 
        }
    }
}
