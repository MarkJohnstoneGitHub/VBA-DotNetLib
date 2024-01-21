// https://learn.microsoft.com/en-us/dotnet/api/system.urihostnametype?view=netframework-4.8.1
// https://referencesource.microsoft.com/#System/net/System/UriHostNameType.cs

using System;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("21E0513B-D040-4C6C-9290-E3595D554982")]

    public enum UriHostNameType
    {
        /// <devdoc>
        ///    <para>[To be supplied.]</para>
        /// </devdoc>
        Unknown,
        /// <devdoc>
        ///    <para>[To be supplied.]</para>
        /// </devdoc>
        Basic,
        /// <devdoc>
        ///    <para>[To be supplied.]</para>
        /// </devdoc>
        Dns,
        /// <devdoc>
        ///    <para>[To be supplied.]</para>
        /// </devdoc>
        IPv4,
        /// <devdoc>
        ///    <para>[To be supplied.]</para>
        /// </devdoc>
        IPv6
    }

}
