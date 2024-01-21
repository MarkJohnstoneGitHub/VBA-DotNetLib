// https://learn.microsoft.com/en-us/dotnet/api/system.text.unicodeencoding?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Guid("F3CFEEF0-3665-4CAD-BFCB-66827A048DE8")]
    [Description("Represents a UTF-16 encoding of Unicode characters.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IUnicodeEncodingSingleton
    {
        // Factory Methods

        //[Description("Initializes a new instance of the UnicodeEncoding class")]
        //UnicodeEncoding Create2();

        //[Description("Initializes a new instance of the UnicodeEncoding class. Parameters specify whether to use the big endian byte order and whether the GetPreamble() method returns a Unicode byte order mark.")]
        //UnicodeEncoding Create2(bool bigEndian, bool byteOrderMark);

        [Description("Initializes a new instance of the UnicodeEncoding class. Parameters specify whether to use the big endian byte order, whether to provide a Unicode byte order mark, and whether to throw an exception when an invalid encoding is detected.")]
        UnicodeEncoding Create(bool bigEndian = false, bool byteOrderMark = true, bool throwOnInvalidBytes = false);

        // Fields
        int CharSize
        {
            [Description("Represents the Unicode character size in bytes. This field is a constant.")]
            get; 
        }

    }
}
