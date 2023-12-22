// https://learn.microsoft.com/en-us/dotnet/api/system.io.streamwriter?view=netframework-4.8.1

using GSystem = global::System;
using GIO = global::System.IO;
using GText = global::System.Text;
using GRemoting = global::System.Runtime.Remoting;
using GTasks = global::System.Threading.Tasks;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("F0DBA134-1C69-4C41-8C72-47709BAE37F2")]
    [Description("Implements a TextWriter for writing characters to a stream in a particular encoding.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStreamWriter
    {
        // Properties
        bool AutoFlush
        {
            [Description("Gets or sets a value indicating whether the StreamWriter will flush its buffer to the underlying stream after every call to Write(Char).")]
            get;
            [Description("Gets or sets a value indicating whether the StreamWriter will flush its buffer to the underlying stream after every call to Write(Char).")]
            set;
        }
        GIO.Stream BaseStream
        {
            [Description("Gets the underlying stream that interfaces with a backing store.")]
            get;
        }

        GText.Encoding Encoding
        {
            [Description("Gets the Encoding in which the output is written.")]
            get;
        }

        IFormatProvider FormatProvider
        {
            [Description("Gets an object that controls formatting.\r\n\r\n(Inherited from TextWriter)")]
            get;
        }

        string NewLine
        {
            [Description("Gets or sets the line terminator string used by the current TextWriter.\r\n\r\n(Inherited from TextWriter)")]
            get;
            [Description("Gets or sets the line terminator string used by the current TextWriter.\r\n\r\n(Inherited from TextWriter)")]
            set;
        }

        // Methods

        [Description("Closes the current StreamWriter object and the underlying stream.")]
        void Close();

        [Description("Creates an object that contains all the relevant information required to generate a proxy used to communicate with a remote object.\r\n\r\n(Inherited from MarshalByRefObject)")]
        GRemoting.ObjRef CreateObjRef(Type requestedType);

        [Description("Releases all resources used by the TextWriter object.\r\n\r\n(Inherited from TextWriter)")]
        void Dispose();

        [Description("Determines whether the specified object is equal to the current object.\r\n\r\n(Inherited from Object)")]
        bool Equals(object obj);

        [Description("Clears all buffers for the current writer and causes any buffered data to be written to the underlying stream.")]
        void Flush();

        [Description("Clears all buffers for this stream asynchronously and causes any buffered data to be written to the underlying device.")]
        GTasks.Task FlushAsync();

        [Description("Serves as the default hash function.\r\n\r\n(Inherited from Object)")]
        int GetHashCode();

        [Description("Retrieves the current lifetime service object that controls the lifetime policy for this instance.\r\n\r\n(Inherited from MarshalByRefObject)")]
        object GetLifetimeService();

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

        [Description("Obtains a lifetime service object to control the lifetime policy for this instance.\r\n\r\n(Inherited from MarshalByRefObject)")] 
        object InitializeLifetimeService();

        [Description("Returns a string that represents the current object.\r\n\r\n(Inherited from Object)")]
        string ToString();

        [Description("Writes a string to the stream.")]
        void Write(string value);

        [Description("Writes the text representation of a Boolean value to the text stream.\r\n\r\n(Inherited from TextWriter)")]
        void Write2(bool value);

        [Description("Writes the text representation of a 4-byte signed integer to the text stream.\r\n\r\n(Inherited from TextWriter)")]
        void Write3(int value);

        [Description("Writes the text representation of an 8-byte signed integer to the text stream.\r\n\r\n(Inherited from TextWriter)")]
        void Write4(long value);

        [Description("Writes the text representation of a 4-byte floating-point value to the text stream.\r\n\r\n(Inherited from TextWriter)")]
        void Write5(float value);

        [Description("Writes the text representation of an 8-byte floating-point value to the text stream.\r\n\r\n(Inherited from TextWriter)")]
        void Write6(double value);

        [Description("Writes the text representation of an object to the text stream by calling the ToString method on that object.\r\n\r\n(Inherited from TextWriter)")]
        void Write7(object value);

        [Description("")]
        void Write8(string format, object arg0);

        [Description("")]
        void Write9(string format, object arg0, object arg1);

        [Description("")]
        void Write10(string format, object arg0, object arg1, object arg2);

        [Description("Writes a formatted string to the text stream, using the same semantics as the Format(String, Object[]) method.\r\n\r\n(Inherited from TextWriter)")]
        void Write11(string format, [In] ref object[] arg);

        [Description("Writes a line terminator to the text stream.\r\n\r\n(Inherited from TextWriter)")]
        void WriteLine();

        [Description("Writes a line terminator to the text stream.\r\n\r\n(Inherited from TextWriter)")]
        void WriteLine2(string value);

        [Description("Writes the text representation of a Boolean value to the text stream, followed by a line terminator.\r\n\r\n(Inherited from TextWriter)")]
        void WriteLine3(bool value);

        [Description("Writes the text representation of a 4-byte signed integer to the text stream, followed by a line terminator.\r\n\r\n(Inherited from TextWriter)")]
        void WriteLine4(int value);

        [Description("Writes the text representation of an 8-byte signed integer to the text stream, followed by a line terminator.\r\n\r\n(Inherited from TextWriter)")]
        void WriteLine5(long value);

        [Description("Writes the text representation of a 4-byte floating-point value to the text stream, followed by a line terminator.\r\n\r\n(Inherited from TextWriter)")]
        void WriteLine6(float value);

        [Description("Writes the text representation of a 8-byte floating-point value to the text stream, followed by a line terminator.\r\n\r\n(Inherited from TextWriter)")]
        void WriteLine7(double value);

        [Description("Writes the text representation of an object to the text stream, by calling the ToString method on that object, followed by a line terminator.\r\n\r\n(Inherited from TextWriter)")]
        void WriteLine8(object value);

        [Description("Asynchronously writes a line terminator to the stream.")]
        GTasks.Task WriteLineAsync();

        [Description("Asynchronously writes a line terminator to the stream.")]
        GTasks.Task WriteLineAsync(string value);
    }
}
