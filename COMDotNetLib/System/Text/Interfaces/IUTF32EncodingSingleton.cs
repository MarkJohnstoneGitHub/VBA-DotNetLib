// https://learn.microsoft.com/en-us/dotnet/api/system.text.utf32encoding?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Guid("901CEEA1-8DEC-48F7-B1FF-0E1897ECEA71")]
    [Description("Represents a UTF-32 encoding of Unicode characters.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IUTF32EncodingSingleton
    {
        [Description("Initializes a new instance of the UTF32Encoding class. Parameters specify whether to use the big endian byte order, whether to provide a Unicode byte order mark, and whether to throw an exception when an invalid encoding is detected.")]
        UTF32Encoding Create(bool bigEndian = false, bool byteOrderMark = true, bool throwOnInvalidCharacters = false);
    }
}
