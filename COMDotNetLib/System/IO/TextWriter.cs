// https://learn.microsoft.com/en-us/dotnet/api/system.io.textwriter?view=netframework-4.8.1


using GSystem = global::System;
using GIO = global::System.IO;
using GText = global::System.Text;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    public interface TextWriter
    {
        Encoding Encoding 
        {
            [Description("")]
            get;
        }

        IFormatProvider FormatProvider 
        {
            [Description("")]
            get;
        }

        string NewLine 
        {
            [Description("")]
            get;
            [Description("")]
            set;
        }

        [Description("")]
        void Close();

        [Description("")]
        void Dispose();

        [Description("")]
        void Flush();

        [Description("Writes a string to the text stream.")]
        void Write(string value);

        [Description("Writes the text representation of a Boolean value to the text stream.")]
        void Write2(bool value);

        [Description("Writes the text representation of a 4-byte signed integer to the text stream.")]
        void Write3(int value);

        [Description("Writes the text representation of an 8-byte signed integer to the text stream.")]
        void Write4(long value);

        [Description("Writes the text representation of a 4-byte floating-point value to the text stream.")]
        void Write5(float value);

        [Description("Writes the text representation of an 8-byte floating-point value to the text stream.")]
        void Write6(double value);

        [Description("Writes the text representation of an object to the text stream by calling the ToString method on that object.")]
        void Write7(object value);

        [Description("")]
        void Write8(string format, object arg0);

        [Description("")]
        void Write9(string format, object arg0, object arg1);

        [Description("")]
        void Write10(string format, object arg0, object arg1, object arg2);

        [Description("Writes a formatted string to the text stream, using the same semantics as the Format(String, Object[]) method.")]
        void Write11(string format, [In] ref object[] arg);

        void Write(byte[] buffer, int index, int count);




        [Description("Writes a line terminator to the text stream.")]
        void WriteLine();

        [Description("Writes a line terminator to the text stream.")]
        void WriteLine2(string value);

        [Description("Writes the text representation of a Boolean value to the text stream, followed by a line terminator.")]
        void WriteLine3(bool value);

        [Description("Writes the text representation of a 4-byte signed integer to the text stream, followed by a line terminator.")]
        void WriteLine4(int value);

        [Description("Writes the text representation of an 8-byte signed integer to the text stream, followed by a line terminator.")]
        void WriteLine5(long value);

        [Description("Writes the text representation of a 4-byte floating-point value to the text stream, followed by a line terminator.")]
        void WriteLine6(float value);

        [Description("Writes the text representation of a 8-byte floating-point value to the text stream, followed by a line terminator.")]
        void WriteLine7(double value);

        [Description("Writes the text representation of an object to the text stream, by calling the ToString method on that object, followed by a line terminator.")]
        void WriteLine8(object value);

    }
}
