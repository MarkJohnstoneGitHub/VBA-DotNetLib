using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("C6E48F40-EAC0-49D8-BF41-3D2F28BA1B72")]
    [Description("Represents the standard input, output, and error streams for console applications.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IConsoleSingleton
    {
        [Description("Clears the console buffer and corresponding console window of display information.")]
        void Clear();

        [Description("Writes the current line terminator to the standard output stream.")]
        void WriteLine();

        [Description("Writes the specified string value, followed by the current line terminator, to the standard output stream.")]
        void WriteLine2(string value);

        [Description("Writes the text representation of the specified object, followed by the current line terminator, to the standard output stream using the specified format information.")]
        void WriteLine3(string format, object arg0);

        [Description("Writes the text representation of the specified object, followed by the current line terminator, to the standard output stream using the specified format information.")]
        void WriteLine4(string format, object arg0, object arg1);

        [Description("Writes the text representation of the specified object, followed by the current line terminator, to the standard output stream using the specified format information.")]
        void WriteLine5(string format, object arg0, object arg1, object arg2);

        [Description("Writes the text representation of the specified object, followed by the current line terminator, to the standard output stream using the specified format information.")]
        void WriteLine6(string format, object arg0, object arg1, object arg2, object arg3);

        [Description("Writes the text representation of the specified array of objects, followed by the current line terminator, to the standard output stream using the specified format information.")]
        void WriteLine7(string format, [In] ref object[] arg);

    }
}
