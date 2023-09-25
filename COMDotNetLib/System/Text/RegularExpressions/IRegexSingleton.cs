// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("8D42ED1E-623C-4704-B008-4AEDBE5A36FD")]
    [Description("Represents an immutable regular expression.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRegexSingleton
    {
        [Description("Escapes a minimal set of characters (\\, *, +, ?, |, {, [, (,), ^, $, ., #, and white space) by replacing them with their escape codes. This instructs the regular expression engine to interpret these characters literally rather than as metacharacters.")]
        string Escape(string str);

        [Description("Converts any escaped characters in the input string.")]
        string Unescape(string str);

    }
}
