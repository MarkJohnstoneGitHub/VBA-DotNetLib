// // https://learn.microsoft.com/en-us/dotnet/api/system.uri?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;
using GSystem = global::System;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("BC8C8386-337E-455C-ACF9-B034C406DE75")]
    [ProgId("DotNetLib.System.UriSingleton")]
    [Description("Provides an object representation of a uniform resource identifier (URI) and easy access to the parts of the URI.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IUriSingleton))]
    public class UriSingleton : IUriSingleton
    {
        public UriSingleton() { }

        // Factory Methods
        public Uri Create(string uriString)
        {
            return new Uri(uriString);
        }

        public Uri Create2(Uri baseUri, string relativeUri)
        {
            return new Uri(baseUri, relativeUri);
        }

        public Uri Create3(Uri baseUri, Uri relativeUri)
        {
            return new Uri(baseUri, relativeUri);
        }

        public Uri Create4(string uriString, UriKind uriKind)
        {
            return new Uri(uriString, uriKind);
        }



        // Fields
        public string SchemeDelimiter => GSystem.Uri.SchemeDelimiter;

        public string UriSchemeFile => GSystem.Uri.UriSchemeFile;

        public string UriSchemeFtp => GSystem.Uri.UriSchemeFtp;

        public  string UriSchemeGopher => GSystem.Uri.UriSchemeGopher;

        public string UriSchemeHttp => GSystem.Uri.UriSchemeHttp;

        public string UriSchemeHttps => GSystem.Uri.UriSchemeHttps;

        public string UriSchemeMailto => GSystem.Uri.UriSchemeMailto;

        public string UriSchemeNetPipe => GSystem.Uri.UriSchemeNetPipe;

        public string UriSchemeNetTcp => GSystem.Uri.UriSchemeNetTcp;

        public string UriSchemeNews => GSystem.Uri.UriSchemeNews;

        public string UriSchemeNntp => GSystem.Uri.UriSchemeNntp;

        // Methods
        public UriHostNameType CheckHostName(string name)
        {
            return (UriHostNameType)GSystem.Uri.CheckHostName(name);
        }

        public bool CheckSchemeName(string schemeName)
        {
            return GSystem.Uri.CheckSchemeName(schemeName);
        }

        public int Compare(Uri uri1, Uri uri2, UriComponents partsToCompare, UriFormat compareFormat, GSystem.StringComparison comparisonType)
        {
            return GSystem.Uri.Compare(uri1.WrappedUri, uri2.WrappedUri, (GSystem.UriComponents)partsToCompare, (GSystem.UriFormat)compareFormat, comparisonType);
        }

        public string EscapeDataString(string stringToEscape)
        {
            return GSystem.Uri.EscapeDataString(stringToEscape);
        }

        public string EscapeUriString(string stringToEscape)
        {
            return GSystem.Uri.EscapeUriString(stringToEscape);
        }

        public int FromHex(string digit)
        {
            return GSystem.Uri.FromHex(digit[0]);
        }

        public string HexEscape(string character)
        {
            return GSystem.Uri.HexEscape(character[0]);
        }

        public string HexUnescape(string pattern, ref int index)
        {
            return GSystem.Uri.HexUnescape(pattern, ref index).ToString();
        }

        public bool IsHexDigit(string character)
        {
            return GSystem.Uri.IsHexDigit(character[0]);
        }

        public bool IsHexEncoding(string pattern, int index)
        {
            return GSystem.Uri.IsHexEncoding(pattern, index);
        }

        public bool IsWellFormedUriString(string uriString, UriKind uriKind)
        {
            return GSystem.Uri.IsWellFormedUriString(uriString, (GSystem.UriKind)uriKind);
        }

        public bool TryCreate(Uri baseUri, Uri relativeUri, out Uri result)
        {
            GSystem.Uri outResult;
            bool tryCreate;
            tryCreate = GSystem.Uri.TryCreate(baseUri.WrappedUri, relativeUri.WrappedUri, out outResult);
            result = new Uri(outResult);
            return tryCreate;
        }

        public bool TryCreate2(Uri baseUri, string relativeUri, out Uri result)
        {
            GSystem.Uri outResult;
            bool tryCreate;
            tryCreate = GSystem.Uri.TryCreate(baseUri.WrappedUri, relativeUri, out outResult);
            result = new Uri(outResult);
            return tryCreate;
        }

        public bool TryCreate3(string uriString, UriKind uriKind, out Uri result)
        {
            GSystem.Uri outResult;
            bool tryCreate;
            tryCreate = GSystem.Uri.TryCreate(uriString, (GSystem.UriKind)uriKind, out outResult);
            result = new Uri(outResult);
            return tryCreate;
        }

        public string UnescapeDataString(string stringToUnescape)
        {
            return GSystem.Uri.UnescapeDataString(stringToUnescape);
        }

        // Operators

        public bool Equality(Uri uri1, Uri uri2)
        {
            return uri1 == uri2;
        }

        public bool Inequality(Uri uri1, Uri uri2)
        {
            return uri1 != uri2;
        }
    }
}
