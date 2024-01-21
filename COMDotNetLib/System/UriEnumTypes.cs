// https://referencesource.microsoft.com/#System/net/System/UriEnumTypes.cs

using System;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    // https://learn.microsoft.com/en-us/dotnet/api/system.uriformat?view=netframework-4.8.1

    [ComVisible(true)]
    [Guid("528A23C3-15AF-4309-9B60-F1C90D0ABBFF")]

    //
    // Summary:
    //     Controls how URI information is escaped.
    public enum UriFormat
    {
        //
        // Summary:
        //     Escaping is performed according to the rules in RFC 2396.
        UriEscaped = 1,
        //
        // Summary:
        //     No escaping is performed.
        Unescaped,
        //
        // Summary:
        //     Characters that have a reserved meaning in the requested URI components remain
        //     escaped. All others are not escaped.
        SafeUnescaped
    }

    [ComVisible(true)]
    [Guid("153E505D-FCBF-4E71-BA8D-9280B182C1B4")]
    // Used to control whether absolute
    public enum UriKind
    {
        RelativeOrAbsolute = 0,
        Absolute = 1,
        Relative = 2
    }


    [ComVisible(true)]
    [Guid("8B89BBA4-2B56-4736-B6BF-48C79700D8B0")]

    //
    // Summary:
    //     Specifies the parts of a System.Uri.
    [Flags]
    public enum UriComponents
    {
        //
        // Summary:
        //     The System.Uri.Scheme data.
        Scheme = 1,
        //
        // Summary:
        //     The System.Uri.UserInfo data.
        UserInfo = 2,
        //
        // Summary:
        //     The System.Uri.Host data.
        Host = 4,
        //
        // Summary:
        //     The System.Uri.Port data.
        Port = 8,
        //
        // Summary:
        //     The System.Uri.LocalPath data.
        Path = 0x10,
        //
        // Summary:
        //     The System.Uri.Query data.
        Query = 0x20,
        //
        // Summary:
        //     The System.Uri.Fragment data.
        Fragment = 0x40,
        //
        // Summary:
        //     The System.Uri.Port data. If no port data is in the System.Uri and a default
        //     port has been assigned to the System.Uri.Scheme, the default port is returned.
        //     If there is no default port, -1 is returned.
        StrongPort = 0x80,
        //
        // Summary:
        //     The normalized form of the System.Uri.Host.
        NormalizedHost = 0x100,
        //
        // Summary:
        //     Specifies that the delimiter should be included.
        KeepDelimiter = 0x40000000,
        //
        // Summary:
        //     The complete System.Uri context that is needed for Uri Serializers. The context
        //     includes the IPv6 scope.
        SerializationInfoString = int.MinValue,
        //
        // Summary:
        //     The System.Uri.Scheme, System.Uri.UserInfo, System.Uri.Host, System.Uri.Port,
        //     System.Uri.LocalPath, System.Uri.Query, and System.Uri.Fragment data.
        AbsoluteUri = 0x7F,
        //
        // Summary:
        //     The System.Uri.Host and System.Uri.Port data. If no port data is in the Uri and
        //     a default port has been assigned to the System.Uri.Scheme, the default port is
        //     returned. If there is no default port, -1 is returned.
        HostAndPort = 0x84,
        //
        // Summary:
        //     The System.Uri.UserInfo, System.Uri.Host, and System.Uri.Port data. If no port
        //     data is in the System.Uri and a default port has been assigned to the System.Uri.Scheme,
        //     the default port is returned. If there is no default port, -1 is returned.
        StrongAuthority = 0x86,
        //
        // Summary:
        //     The System.Uri.Scheme, System.Uri.Host, and System.Uri.Port data.
        SchemeAndServer = 0xD,
        //
        // Summary:
        //     The System.Uri.Scheme, System.Uri.Host, System.Uri.Port, System.Uri.LocalPath,
        //     and System.Uri.Query data.
        HttpRequestUrl = 0x3D,
        //
        // Summary:
        //     The System.Uri.LocalPath and System.Uri.Query data. Also see System.Uri.PathAndQuery.
        PathAndQuery = 0x30
    }
}
