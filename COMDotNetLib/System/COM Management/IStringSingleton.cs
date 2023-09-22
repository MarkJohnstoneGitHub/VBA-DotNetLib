// https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("4F39D64D-9BD3-4AF7-A124-7A88364BE29F")]
    [Description("")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStringSingleton
    {
        [Description("Replaces the format item in a specified string with the string representation of a corresponding object in a specified array.")]
        string Format(string pFormat, [In] ref object[] args);

        [Description("Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information.")]
        string Format2(IFormatProvider provider, string pFormat, [In] ref object[] args);
    }
}
