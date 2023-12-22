using GSystem = global::System;
using GIO = global::System.IO;

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("C7D72694-97BA-4888-B71C-B0D362951A5B")]
    [Description("")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface Stream
    {
        bool CanRead 
        {
            [Description("When overridden in a derived class, gets a value indicating whether the current stream supports reading.")]
            get;
        }

        bool CanSeek 
        {
            [Description("When overridden in a derived class, gets a value indicating whether the current stream supports seeking.")]
            get;
        }

        bool CanTimeout 
        {
            [Description("Gets a value that determines whether the current stream can time out.")]
            get;
        }

        bool CanWrite 
        {
            [Description("When overridden in a derived class, gets a value indicating whether the current stream supports writing.")]
            get;
        }

        long Length 
        {
            [Description("When overridden in a derived class, gets the length in bytes of the stream.")]
            get;
        }

        long Position 
        {
            [Description("When overridden in a derived class, gets or sets the position within the current stream.")]
            get;
            [Description("When overridden in a derived class, gets or sets the position within the current stream.")]
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

        [Description("Begins an asynchronous read operation. (Consider using ReadAsync(Byte[], Int32, Int32) instead.)")]
        GSystem.IAsyncResult BeginRead([In] ref byte[] buffer, int offset, int count, GSystem.AsyncCallback callback, object state);

        [Description("Begins an asynchronous write operation. (Consider using WriteAsync(Byte[], Int32, Int32) instead.)")]
        GSystem.IAsyncResult BeginWrite([In] ref byte[] buffer, int offset, int count, GSystem.AsyncCallback callback, object state);

        [Description("Closes the current stream and releases any resources (such as sockets and file handles) associated with the current stream. Instead of calling this method, ensure that the stream is properly disposed.")]
        void Close();

        [Description("Reads the bytes from the current stream and writes them to another stream. Both streams positions are advanced by the number of bytes copied.")]
        void CopyTo(Stream destination);

        [Description("Reads the bytes from the current stream and writes them to another stream, using a specified buffer size. Both streams positions are advanced by the number of bytes copied.")]
        void CopyTo(Stream destination, int bufferSize);

        [Description("Releases all resources used by the Stream.")]
        void Dispose();

        [Description("Waits for the pending asynchronous read to complete. (Consider using ReadAsync(Byte[], Int32, Int32) instead.)")]
        int EndRead(IAsyncResult asyncResult);

        [Description("Ends an asynchronous write operation. (Consider using WriteAsync(Byte[], Int32, Int32) instead.)")]
        void EndWrite(IAsyncResult asyncResult);

        [Description("When overridden in a derived class, clears all buffers for this stream and causes any buffered data to be written to the underlying device.")]
        void Flush();

        [Description("When overridden in a derived class, reads a sequence of bytes from the current stream and advances the position within the stream by the number of bytes read.")]
        int Read([In][Out] ref byte[] buffer, int offset, int count);

        [Description("Reads a byte from the stream and advances the position within the stream by one byte, or returns -1 if at the end of the stream.")]
        int ReadByte();

        [Description("When overridden in a derived class, sets the position within the current stream.")]
        long Seek(long offset, GIO.SeekOrigin origin);

        [Description("When overridden in a derived class, sets the length of the current stream.")]
        void SetLength(long value);

        [Description("When overridden in a derived class, writes a sequence of bytes to the current stream and advances the current position within this stream by the number of bytes written.")]
        void Write([In][Out] ref byte[] buffer, int offset, int count);

        [Description("Writes a byte to the current position in the stream and advances the position within the stream by one byte.")]
        void WriteByte(byte value);
    }
}
