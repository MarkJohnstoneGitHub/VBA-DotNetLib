// https://learn.microsoft.com/en-us/dotnet/api/system.io.filestream?view=netframework-4.8.1

using GSystem = global::System;
using GIO = global::System.IO;
using GThreading = global::System.Threading;
using GTasks = global::System.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("B5662303-7652-415D-BE70-46C10540B2DD")]
    [Description("Provides a Stream for a file, supporting both synchronous and asynchronous read and write operations.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IFileStream
    {
        // Properties
        bool CanRead
        {
            [Description("Gets a value indicating whether the current stream supports reading.")]
            get;
        }

        bool CanSeek
        {
            [Description("Gets a value indicating whether the current stream supports seeking.")]
            get;
        }

        bool CanTimeout
        {
            [Description("Gets a value that determines whether the current stream can time out.\r\n\r\n(Inherited from Stream)")]
            get;
        }

        bool CanWrite
        {
            [Description("Gets a value indicating whether the current stream supports writing.")]
            get;
        }

        bool IsAsync 
        {
            [Description("Gets a value that indicates whether the FileStream was opened asynchronously or synchronously.")]
            get;
        }

        long Length
        {
            [Description("Gets the stream length in bytes.")]
            get;
        }

        string Name 
        {
            [Description("Gets the absolute path of the file opened in the FileStream.")]
            get;
        }

        long Position
        {
            [Description("Gets the position within the current stream.")]
            get;
            [Description("Gets the position within the current stream.")]
            set;
        }

        int ReadTimeout 
        {
            [Description("Gets or sets a value, in milliseconds, that determines how long the stream will attempt to read before timing out.")]
            get;
            [Description("Gets or sets a value, in milliseconds, that determines how long the stream will attempt to read before timing out.")]
            set;
        }

        int WriteTimeout 
        {
            [Description("Gets or sets a value, in milliseconds, that determines how long the stream will attempt to write before timing out.")]
            get;
            [Description("Gets or sets a value, in milliseconds, that determines how long the stream will attempt to write before timing out.")]
            set;
        }

        // Methods

        // Methods

        [Description("Begins an asynchronous read operation. (Consider using ReadAsync(Byte[], Int32, Int32, CancellationToken) instead.)")]
        GSystem.IAsyncResult BeginRead([In] ref byte[] buffer, int offset, int count, GSystem.AsyncCallback callback, object state);

        [Description("Begins an asynchronous write operation. (Consider using WriteAsync(Byte[], Int32, Int32, CancellationToken) instead.)")]
        GSystem.IAsyncResult BeginWrite([In] ref byte[] buffer, int offset, int count, GSystem.AsyncCallback callback, object state);

        [Description("Closes the current stream and releases any resources (such as sockets and file handles) associated with the current stream. Instead of calling this method, ensure that the stream is properly disposed.\r\n\r\n(Inherited from Stream)")]
        void Close();

        [Description("Reads the bytes from the current stream and writes them to another stream. Both streams positions are advanced by the number of bytes copied.\r\n\r\n(Inherited from Stream)")]
        void CopyTo(Stream destination);

        [Description("Reads the bytes from the current buffered stream and writes them to another stream.")]
        void CopyTo(Stream destination, int bufferSize);

        [Description("Releases all resources used by the Stream.\r\n\r\n(Inherited from Stream)")]
        void Dispose();

        [Description("Waits for the pending asynchronous read operation to complete. (Consider using ReadAsync(Byte[], Int32, Int32, CancellationToken) instead.)")]
        int EndRead(IAsyncResult asyncResult);

        [Description("Ends an asynchronous write operation and blocks until the I/O operation is complete. (Consider using WriteAsync(Byte[], Int32, Int32, CancellationToken) instead.)")]
        void EndWrite(IAsyncResult asyncResult);

        [Description("Determines whether the specified object is equal to the current object.\r\n\r\n(Inherited from Object)")]
        bool Equals(object obj);

        [Description("Clears all buffers for this stream and causes any buffered data to be written to the underlying device.")]
        void Flush();

        //[Description("Asynchronously clears all buffers for this stream and causes any buffered data to be written to the underlying device.\r\n\r\n(Inherited from Stream)")]
        //GTasks.Task FlushAsync(GThreading.CancellationToken cancellationToken);

        [Description("Serves as the default hash function.\r\n\r\n(Inherited from Object)")]
        int GetHashCode();

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

        [Description("Copies bytes from the current buffered stream to an array.")]

        int Read([In][Out] ref byte[] buffer, int offset, int count);

        [Description("Reads a byte from the underlying stream and returns the byte cast to an int, or returns -1 if reading from the end of the stream.")]
        int ReadByte();

        [Description("Sets the position within the current buffered stream.")]
        long Seek(long offset, GIO.SeekOrigin origin);

        [Description("Sets the length of the buffered stream.")]
        void SetLength(long value);

        [Description("Returns a string that represents the current object.\r\n\r\n(Inherited from Object)")]
        string ToString();

        [Description("Copies bytes to the buffered stream and advances the current position within the buffered stream by the number of bytes written.")]
        void Write([In][Out] ref byte[] buffer, int offset, int count);

        //[Description("Asynchronously writes a sequence of bytes to the current stream and advances the current position within this stream by the number of bytes written.\r\n\r\n(Inherited from Stream)")]
        //GTasks.Task WriteAsync(byte[] buffer, int offset, int count, GThreading.CancellationToken cancellationToken);

        [Description("Writes a byte to the current position in the buffered stream.")]
        void WriteByte(byte value);


        // [Description("")]

    }
}
