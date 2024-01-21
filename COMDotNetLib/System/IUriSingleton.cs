// https://learn.microsoft.com/en-us/dotnet/api/system.uri?view=netframework-4.8.1

using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("BEF6D2F2-31FE-42E3-B19F-E414991E763D")]
    [Description("Provides an object representation of a uniform resource identifier (URI) and easy access to the parts of the URI.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IUriSingleton
    {
        // Factory Methods

        [Description("Initializes a new instance of the Uri class with the specified URI.")]
        Uri Create(string uriString);

        [Description("Initializes a new instance of the Uri class with the specified URI. This constructor allows you to specify if the URI string is a relative URI, absolute URI, or is indeterminate.")]
        Uri Create4(string uriString, UriKind uriKind);

        [Description("Initializes a new instance of the Uri class based on the specified base URI and relative URI string.")]
        Uri Create2(Uri baseUri, string relativeUri);

        [Description("Initializes a new instance of the Uri class based on the combination of a specified base Uri instance and a relative Uri instance.")]
        Uri Create3(Uri baseUri, Uri relativeUri);

        // Fields

        string SchemeDelimiter
        {
            [Description("Specifies the characters that separate the communication protocol scheme from the address portion of the URI. This field is read-only.")]
            get;
        }

        string UriSchemeFile
        {
            [Description("Specifies that the URI is a pointer to a file. This field is read-only.")]
            get;
        }

        string UriSchemeFtp
        {
            [Description("Specifies that the URI is accessed through the File Transfer Protocol (FTP). This field is read-only.")]
            get;
        }

        string UriSchemeGopher
        {
            [Description("Specifies that the URI is accessed through the Gopher protocol. This field is read-only.")]
            get;
        }

        string UriSchemeHttp
        {
            [Description("Specifies that the URI is accessed through the Hypertext Transfer Protocol (HTTP). This field is read-only.")]
            get;
        }

        string UriSchemeHttps
        {
            [Description("Specifies that the URI is accessed through the Secure Hypertext Transfer Protocol (HTTPS). This field is read-only.")]
            get;
        }

        string UriSchemeMailto
        {
            [Description("Specifies that the URI is an email address and is accessed through the Simple Mail Transport Protocol (SMTP). This field is read-only.")]
            get;
        }

        string UriSchemeNetPipe
        {
            [Description("Specifies that the URI is accessed through the NetPipe scheme used by Windows Communication Foundation (WCF). This field is read-only.")]
            get;
        }

        string UriSchemeNetTcp
        {
            [Description("Specifies that the URI is accessed through the NetTcp scheme used by Windows Communication Foundation (WCF). This field is read-only.")]
            get;
        }

        string UriSchemeNews
        {
            [Description("Specifies that the URI is an Internet news group and is accessed through the Network News Transport Protocol (NNTP). This field is read-only.")]
            get;
        }

        string UriSchemeNntp
        {
            [Description("Specifies that the URI is an Internet news group and is accessed through the Network News Transport Protocol (NNTP). This field is read-only.")]
            get;
        }

        // Methods

        [Description("Determines whether the specified host name is a valid DNS name.")]
        UriHostNameType CheckHostName(string name);

        [Description("Determines whether the specified scheme name is valid.")]
        bool CheckSchemeName(string schemeName);

        [Description("Compares the specified parts of two URIs using the specified comparison rules.")]
        int Compare(Uri uri1, Uri uri2, UriComponents partsToCompare, UriFormat compareFormat, GSystem.StringComparison comparisonType);

        [Description("Converts a string to its escaped representation.")]
        string EscapeDataString(string stringToEscape);

        [Description("Converts a URI string to its escaped representation.")]
        string EscapeUriString(string stringToEscape);

        [Description("Gets the decimal value of a hexadecimal digit.")]
        int FromHex(string digit);

        [Description("Converts a specified character into its hexadecimal equivalent.")]
        string HexEscape(string character);

        [Description("Converts a specified hexadecimal representation of a character to the character")]
        string HexUnescape(string pattern, ref int index);

        [Description("Determines whether a specified character is a valid hexadecimal digit.")]
        bool IsHexDigit(string character);

        [Description("Determines whether a character in a string is hexadecimal encoded.")]
        bool IsHexEncoding(string pattern, int index);

        [Description("Indicates whether the string is well-formed by attempting to construct a URI with the string and ensures that the string does not require further escaping.")]
        bool IsWellFormedUriString(string uriString, UriKind uriKind);

        [Description("Creates a new Uri using the specified base and relative Uri instances.")]
        bool TryCreate(Uri baseUri, Uri relativeUri, out Uri result);

        [Description("Creates a new Uri using the specified String instance and a UriKind.")]
        bool TryCreate3(string uriString, UriKind uriKind, out Uri result);

        [Description("Creates a new Uri using the specified base and relative String instances.")]
        bool TryCreate2(Uri baseUri, string relativeUri, out Uri result);

        [Description("Converts a string to its unescaped representation.")]
        string UnescapeDataString(string stringToUnescape);

        // Operators

        [Description("Determines whether two Uri instances have the same value.")]
        bool Equality(Uri uri1, Uri uri2);

        [Description("Determines whether two Uri instances do not have the same value.")]
        bool Inequality(Uri uri1, Uri uri2);

    }
}
