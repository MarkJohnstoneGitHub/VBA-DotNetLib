// https://learn.microsoft.com/en-us/dotnet/api/system.io.textreader?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Guid("963606C5-D978-478E-B631-9C6C3FFD48DE")]
    [Description("Represents a reader that can read a sequential series of characters.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface TextReader
    {
        [Description("Closes the TextReader and releases any system resources associated with the TextReader.")]
        void Close();

        // GRemoting.ObjRef CreateObjRef(Type requestedType);

        [Description("Releases all resources used by the TextReader object.")]
        void Dispose();

        //[Description("Releases the unmanaged resources used by the TextReader and optionally releases the managed resources.")]
        //void Dispose(bool disposing);

        [Description("Determines whether the specified object is equal to the current object.\r\n\r\n(Inherited from Object)")]
        bool Equals(object obj);

        [Description("Serves as the default hash function.\r\n\r\n(Inherited from Object)")]
        int GetHashCode();

        //[Description("nRetrieves the current lifetime service object that controls the lifetime policy for this instance.\r\n\r\n(Inherited from MarshalByRefObject)")]
        //object GetLifetimeService();

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

        //[Description("Obtains a lifetime service object to control the lifetime policy for this instance.\r\n\r\n(Inherited from MarshalByRefObject)")]
        //object InitializeLifetimeService();

        [Description("Reads the next character without changing the state of the reader or the character source. Returns the next available character without actually reading it from the reader.")]
        int Peek();

        [Description("Reads the next character from the text reader and advances the character position by one character.")]
        int Read();

        //int Read (char[] buffer, int index, int count);
        [Description("Reads a specified maximum number of characters from the current reader and writes the data to a buffer, beginning at the specified index.")]
        int Read([In][Out] ref byte[] buffer, int index, int count);

        //int ReadBlock (char[] buffer, int index, int count);
        [Description("Reads a specified maximum number of characters from the current text reader and writes the data to a buffer, beginning at the specified index.")]
        int ReadBlock([In][Out] ref byte[] buffer, int index, int count);

        [Description("Reads a line of characters from the text reader and returns the data as a string.")]
        string ReadLine();

        [Description("Reads all characters from the current position to the end of the text reader and returns them as one string.")]
        string ReadToEnd();


    }
}
