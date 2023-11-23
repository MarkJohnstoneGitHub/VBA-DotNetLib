// https://learn.microsoft.com/en-us/dotnet/api/system.text.utf7encoding?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Guid("10149089-AE74-4E35-829E-367E66A6FBE4")]
    [Description("Represents a UTF-7 encoding of Unicode characters.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IUTF7EncodingSingleton
    {
        //[Description("Initializes a new instance of the UTF7Encoding class.")]
        //UTF7Encoding Create();

        [Description("Initializes a new instance of the UTF7Encoding class. A parameter specifies whether to allow optional characters.")]
        UTF7Encoding Create(bool allowOptionals = false);
    }
}
