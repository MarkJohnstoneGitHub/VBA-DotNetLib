// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regexoptions?view=netframework-4.8.1

using System;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("19FEC869-C5E3-4986-BCF2-1F8898D12026")]
    public enum RegexOptions
    {
        None = 0x0,
        IgnoreCase = 0x1,
        Multiline = 0x2,
        ExplicitCapture = 0x4,
        Compiled = 0x8,
        Singleline = 0x10,
        IgnorePatternWhitespace = 0x20,
        RightToLeft = 0x40,
        ECMAScript = 0x100,
        CultureInvariant = 0x200,
    }
}
