// https://learn.microsoft.com/en-us/dotnet/api/system.io.filestream?view=netframework-4.8.1

using GSystem = global::System;
using GIO = global::System.IO;
using GThreading = global::System.Threading;
using GTasks = global::System.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Description("Provides a Stream for a file, supporting both synchronous and asynchronous read and write operations.")]
    [Guid("7CE4C6D7-9BFF-4137-A8D4-859A6590F261")]
    [ProgId("DotNetLib.System.IO.FileStream")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IFileStream))]

    public class FileStream : GIO.Stream, IFileStream
    {
        private GIO.FileStream _fileStream;

        public override bool CanRead => _fileStream.CanRead;

        public override bool CanSeek => _fileStream.CanSeek;

        public override bool CanWrite => _fileStream.CanWrite;

        public virtual bool IsAsync => _fileStream.IsAsync;

        public override long Length => _fileStream.Length;

        public string Name => _fileStream.Name;

        public override long Position 
        { 
            get => _fileStream.Position;
            set => _fileStream.Position = value;
        }

        public IAsyncResult BeginRead([In] ref byte[] buffer, int offset, int count, AsyncCallback callback, object state)
        {
            throw new NotImplementedException();
        }

        public IAsyncResult BeginWrite([In] ref byte[] buffer, int offset, int count, AsyncCallback callback, object state)
        {
            throw new NotImplementedException();
        }

        public void CopyTo(Stream destination)
        {
            throw new NotImplementedException();
        }

        public void CopyTo(Stream destination, int bufferSize)
        {
            throw new NotImplementedException();
        }

        public override void Flush()
        {
            _fileStream.Flush();
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return _fileStream.Read(buffer, offset, count);
        }

        public int Read([In][Out] ref byte[] buffer, int offset, int count)
        {
            return _fileStream.Read(buffer, offset, count);
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            return _fileStream.Seek(offset, origin);
        }

        public override void SetLength(long value)
        {
            _fileStream.SetLength(value);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            _fileStream.Write(buffer, offset, count);
        }

        public void Write([In][Out] ref byte[] buffer, int offset, int count)
        {
            _fileStream.Write(buffer, offset, count);
        }

        public new Type GetType()
        {
            throw new NotImplementedException();
        }
    }
}
