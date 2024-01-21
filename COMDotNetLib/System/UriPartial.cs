using System;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("D46C0A36-761F-4E16-B6D3-A18229DB5A47")]

    //
    // Summary:
    //     Defines the parts of a URI for the System.Uri.GetLeftPart(System.UriPartial)
    //     method.
    public enum UriPartial
    {
        //
        // Summary:
        //     The scheme segment of the URI.
        Scheme,
        //
        // Summary:
        //     The scheme and authority segments of the URI.
        Authority,
        //
        // Summary:
        //     The scheme, authority, and path segments of the URI.
        Path,
        //
        // Summary:
        //     The scheme, authority, path, and query segments of the URI.
        Query
    }
}
