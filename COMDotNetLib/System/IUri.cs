// https://learn.microsoft.com/en-us/dotnet/api/system.uri?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("DD4A7552-9F98-4AAC-B120-AD9FCA98343D")]
    [Description("Provides an object representation of a uniform resource identifier (URI) and easy access to the parts of the URI.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IUri
    {

        string AbsolutePath 
        {
            [Description("Gets the absolute path of the URI.")]
            get;
        }

        string AbsoluteUri 
        {
            [Description("Gets the absolute URI.")]
            get;
        }

        string Authority 
        {
            [Description("Gets the Domain Name System (DNS) host name or IP address and the port number for a server.")]
            get;
        }

        string DnsSafeHost 
        {
            [Description("Gets a host name that, after being unescaped if necessary, is safe to use for DNS resolution.")]
            get;
        }

        string Fragment 
        {
            [Description("Gets the escaped URI fragment, including the leading '#' character if not empty.")]
            get;
        }

        string Host 
        {
            [Description("Gets the host component of this instance.")]
            get;
        }

        UriHostNameType HostNameType 
        {
            [Description("Gets the type of the host name specified in the URI.")]
            get;
        }

        string IdnHost 
        {
            [Description("Gets the RFC 3490 compliant International Domain Name of the host, using Punycode as appropriate. This string, after being unescaped if necessary, is safe to use for DNS resolution.")]
            get;
        }

        bool IsAbsoluteUri 
        {
            [Description("Gets a value that indicates whether the Uri instance is absolute.")]
            get;
        }

        bool IsDefaultPort 
        {
            [Description("Gets a value that indicates whether the port value of the URI is the default for this scheme.")]
            get;
        }

        bool IsFile 
        {
            [Description("Gets a value that indicates whether the specified Uri is a file URI.")]
            get;
        }

        bool IsLoopback 
        {
            [Description("Gets a value that indicates whether the specified Uri references the local host.")]
            get;
        }

        bool IsUnc 
        {
            [Description("Gets a value that indicates whether the specified Uri is a universal naming convention (UNC) path.")]
            get;
        }

        string LocalPath 
        {
            [Description("Gets a local operating-system representation of a file name.")]
            get;
        }

        string OriginalString 
        {
            [Description("Gets the original URI string that was passed to the Uri constructor.")]
            get;
        }

        string PathAndQuery 
        {
            [Description("Gets the AbsolutePath and Query properties separated by a question mark (?).")]
            get;
        }

        int Port 
        {
            [Description("Gets the port number of this URI.")]
            get;
        }

        string Query 
        {
            [Description("Gets any query information included in the specified URI, including the leading '?' character if not empty.")]
            get;
        }

        string Scheme 
        {
            [Description("Gets the scheme name for this URI.")]
            get;
        }

        string[] Segments
        {
            [Description("Gets an array containing the path segments that make up the specified URI.")]
            get;
        }

        bool UserEscaped 
        {
            [Description("Gets a value that indicates whether the URI string was completely escaped before the Uri instance was created.")]
            get;
        }

        string UserInfo 
        {
            [Description("Gets the user name, password, or other user-specific information associated with the specified URI.")]
            get;
        }

        // Methods

        [Description("Compares two Uri instances for equality.")]
        bool Equals(object comparand);

        [Description("Gets the specified components of the current instance using the specified escaping for special characters.")]
        string GetComponents(UriComponents components, UriFormat format);

        [Description("Gets the hash code for the URI.")]
        int GetHashCode();

        [Description("Gets the specified portion of a Uri instance.")]
        string GetLeftPart(UriPartial part);

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

        [Description("Determines whether the current Uri instance is a base of the specified Uri instance.")]
        bool IsBaseOf(Uri uri);

        [Description("Indicates whether the string used to construct this Uri was well-formed and does not require further escaping.")]
        bool IsWellFormedOriginalString();

        [Description("Determines the difference between two Uri instances.")]
        Uri MakeRelativeUri(Uri uri);

        [Description("Gets a canonical string representation for the specified Uri instance.")]
        string ToString();


    }
}
