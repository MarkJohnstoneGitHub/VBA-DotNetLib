// https://learn.microsoft.com/en-us/dotnet/api/system.io.bufferedstream?view=netframework-4.8.1

using GSystem = global::System;
using GIO = global::System.IO;
using GThreading = global::System.Threading;
using GTasks = global::System.Threading.Tasks;
using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using DotNetLib.Extensions;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Description("Adds a buffering layer to read and write operations on another stream. This class cannot be inherited.")]
    [Guid("955555E3-9328-4346-BA79-2F761D9C54C5")]
    [ProgId("DotNetLib.System.IO.BufferedStream")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IBufferedStream))]
    public class BufferedStream : GIO.Stream, IBufferedStream, Stream, IDisposable
    {
        private GIO.BufferedStream _bufferedStream;

        // Constructors

        public BufferedStream(GIO.Stream stream)
        {
            _bufferedStream = new GIO.BufferedStream(stream);
        }

        public BufferedStream(GIO.Stream stream, int bufferSize)
        {
            _bufferedStream = new GIO.BufferedStream(stream, bufferSize);
        }

        public BufferedStream(Stream stream)
        {
            _bufferedStream = new GIO.BufferedStream((GIO.Stream)stream);
        }

        public BufferedStream(Stream stream, int bufferSize)
        {
            _bufferedStream = new GIO.BufferedStream((GIO.Stream)stream, bufferSize);
        }

        // Properties

        public override bool CanRead => _bufferedStream.CanRead;

        public override bool CanSeek => _bufferedStream.CanSeek;


        public new bool CanTimeout =>  _bufferedStream.CanTimeout;

        public override bool CanWrite => _bufferedStream.CanWrite;

        public override long Length => _bufferedStream.Length;

        public override long Position 
        { 
            get => _bufferedStream.Position;
            set => _bufferedStream.Position = value;
        }

        public override int ReadTimeout 
        {
            get => _bufferedStream.ReadTimeout;
            set => _bufferedStream.ReadTimeout = value;
        }
        public override int WriteTimeout 
        { 
            get => _bufferedStream.WriteTimeout;
            set => _bufferedStream.WriteTimeout = value;
        }

        // Methods

        public GSystem.IAsyncResult BeginRead([In] ref byte[] buffer, int offset, int count, GSystem.AsyncCallback callback, object state)
        {
            return _bufferedStream.BeginRead(buffer, offset, count, callback, state);
        }

        public GSystem.IAsyncResult BeginWrite([In] ref byte[] buffer, int offset, int count, GSystem.AsyncCallback callback, object state)
        {
            return _bufferedStream.BeginWrite(buffer, offset, count, callback, state);
        }

        public override void Close()
        {
            _bufferedStream.Close();
        }

        // default buffersize 81920 ??
        public void CopyTo(Stream destination)
        {
            _bufferedStream.CopyTo((GIO.Stream)destination);
        }

        public void CopyTo(Stream destination, int bufferSize)
        {
            _bufferedStream.CopyTo((GIO.Stream)destination, bufferSize);
        }

        public new void Dispose()
        {
            _bufferedStream.Dispose();
            //base.Dispose(); //??
        }

        public override int EndRead(IAsyncResult asyncResult)
        {
           return _bufferedStream.EndRead(asyncResult);
        }

        public override void EndWrite(IAsyncResult asyncResult)
        {
            _bufferedStream.EndWrite(asyncResult);
        }

        public override bool Equals(object obj)
        {
            return _bufferedStream.Equals(obj.Unwrap());
        }

        public override void Flush()
        {
            _bufferedStream.Flush();
        }

        public override GTasks.Task FlushAsync(GThreading.CancellationToken cancellationToken)
        { 
            return _bufferedStream.FlushAsync(cancellationToken); 
        }

        public override int GetHashCode()
        { 
            return _bufferedStream.GetHashCode(); 
        }

        public new Type GetType()
        {
            return new Type(((GSystem.Object)this).GetType());
        }

        public int Read([In][Out] ref byte[] buffer, int offset, int count)
        {
            return _bufferedStream.Read(buffer, offset, count);
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return _bufferedStream.Read(buffer, offset, count);
        }

        public override GTasks.Task<int> ReadAsync(byte[] buffer, int offset, int count, GThreading.CancellationToken cancellationToken)
        {
            return _bufferedStream.ReadAsync(buffer, offset, count, cancellationToken);
        }

        public new int ReadByte()
        {
            return _bufferedStream.ReadByte();
        }

        public override long Seek(long offset, GIO.SeekOrigin origin)
        {
            return _bufferedStream.Seek(offset, origin);
        }

        public override void SetLength(long value)
        {
            _bufferedStream.SetLength(value);
        }

        public override string ToString()
        { 
            return _bufferedStream.ToString(); 
        }

        public void Write([In][Out] ref byte[] buffer, int offset, int count)
        {
            _bufferedStream.Write(buffer, offset, count);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            _bufferedStream.Write(buffer, offset, count);
        }

        public override GTasks.Task WriteAsync(byte[] buffer, int offset, int count, GThreading.CancellationToken cancellationToken)
        {
            return _bufferedStream.WriteAsync(buffer, offset, count, cancellationToken);  
        }

        public override void WriteByte(byte value)
        {
            _bufferedStream.WriteByte(value);
        }


    }
}
