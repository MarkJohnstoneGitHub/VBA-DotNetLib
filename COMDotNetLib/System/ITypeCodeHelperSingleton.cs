// https://learn.microsoft.com/en-us/dotnet/api/system.typecode?view=netframework-4.8.1

using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("9A756BE4-62FE-458B-ABCD-A35FA9ABE99B")]
    [Description("")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITypeCodeHelperSingleton
    {
        [Description("Converts the value of this instance to its equivalent string representation.")]
        string ToString(GSystem.TypeCode typecode);

        [Description("Converts the value of this instance to its equivalent string representation using the specified format.")]
        string ToString2(GSystem.TypeCode typecode, string format);
    }
}
