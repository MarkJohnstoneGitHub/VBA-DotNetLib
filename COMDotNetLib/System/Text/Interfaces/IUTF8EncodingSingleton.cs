// https://learn.microsoft.com/en-us/dotnet/api/system.text.utf8encoding?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Guid("EADA552D-9B5E-4273-A05C-27FDDC31D833")]
    [Description("Represents a UTF-8 encoding of Unicode characters.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IUTF8EncodingSingleton
    {
        //[Description("Initializes a new instance of the UTF8Encoding class.")]
        //UTF8Encoding Create();

        //[Description("Initializes a new instance of the UTF8Encoding class. A parameter specifies whether to provide a Unicode byte order mark.")]
        //UTF8Encoding Create(bool encoderShouldEmitUTF8Identifier);

        [Description("Initializes a new instance of the UTF8Encoding class. Parameters specify whether to provide a Unicode byte order mark and whether to throw an exception when an invalid encoding is detected.")]
        UTF8Encoding Create(bool encoderShouldEmitUTF8Identifier = false, bool throwOnInvalidBytes = false);

    }
}
