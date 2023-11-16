using System;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("F8444BF7-BB0B-4987-A5C0-3324B3372092")]

    //
    // Summary:
    //     Specifies whether applicable Overload:System.String.Split method overloads include
    //     or omit empty substrings from the return value.
    [Flags]
    public enum StringSplitOptions
    {
        //
        // Summary:
        //     The return value includes array elements that contain an empty string
        None = 0x0,
        //
        // Summary:
        //     The return value does not include array elements that contain an empty string
        RemoveEmptyEntries = 0x1
    }
}
