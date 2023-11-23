// https://learn.microsoft.com/en-us/dotnet/api/system.text.utf32encoding?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Description("Represents a UTF-32 encoding of Unicode characters.")]
    [Guid("52590749-65B6-401F-AF7D-256DA83BD0EA")]
    [ProgId("DotNetLib.System.Text.UTF32EncodingSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IUTF32EncodingSingleton))]
    public class UTF32EncodingSingleton : IUTF32EncodingSingleton
    {
        public UTF32EncodingSingleton() { }

        public UTF32Encoding Create(bool bigEndian = false, bool byteOrderMark = true, bool throwOnInvalidCharacters = false)
        {
            return new UTF32Encoding(bigEndian, byteOrderMark, throwOnInvalidCharacters);
        }
    }
}
