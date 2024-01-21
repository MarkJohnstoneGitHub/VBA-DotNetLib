// https://learn.microsoft.com/en-us/dotnet/api/system.uri?view=netframework-4.8.1

using GSystem = global::System;
using System;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.Extensions;
using System.Runtime.Serialization;
using System.Security.Permissions;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("B89FB0A9-23ED-40C2-B92B-E788C6644935")]
    [ProgId("DotNetLib.System.Uri")]
    [Description("Provides an object representation of a uniform resource identifier (URI) and easy access to the parts of the URI.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IUri))]
    public class Uri : IUri, IWrappedObject, ISerializable
    {
        private GSystem.Uri _uri;

        internal Uri(GSystem.Uri uri)
        {
            _uri = uri;
        }

        public Uri(string uriString)
        {
            _uri = new GSystem.Uri(uriString);
        }

        public Uri(string uriString, UriKind uriKind)
        {
            _uri = new GSystem.Uri(uriString, (GSystem.UriKind)uriKind);
        }

        public Uri(Uri baseUri, string relativeUri)
        {
            _uri = new GSystem.Uri(baseUri.WrappedUri, relativeUri);
        }

        public Uri(Uri baseUri, Uri relativeUri)
        {
            _uri = new GSystem.Uri(baseUri.WrappedUri, relativeUri.WrappedUri);
        }


        // Properties
        internal GSystem.Uri WrappedUri => _uri;

        public string AbsolutePath => _uri.AbsolutePath;

        public string AbsoluteUri => _uri.AbsoluteUri;

        public string Authority => _uri.Authority;

        public string DnsSafeHost => _uri.DnsSafeHost;

        public string Fragment => _uri.Fragment;

        public string Host => _uri.Host;

        public UriHostNameType HostNameType => (UriHostNameType)_uri.HostNameType;

        public string IdnHost => _uri.IdnHost;

        public bool IsAbsoluteUri => _uri.IsAbsoluteUri;

        public bool IsDefaultPort => _uri.IsDefaultPort;

        public bool IsFile => _uri.IsFile;

        public bool IsLoopback => _uri.IsLoopback;

        public bool IsUnc => _uri.IsUnc;

        public string LocalPath => _uri.LocalPath;

        public string OriginalString => _uri.OriginalString;

        public string PathAndQuery => _uri.PathAndQuery;

        public int Port => _uri.Port;

        public string Query => _uri.Query;

        public string Scheme => _uri.Scheme;

        public string[] Segments => _uri.Segments;

        public bool UserEscaped => _uri.UserEscaped;

        public string UserInfo => _uri.UserInfo;

        public object WrappedObject => _uri;

        // Methods

        public override bool Equals(object comparand)
        {  
            return _uri.Equals(comparand.Unwrap()); 
        }

        public string GetComponents(UriComponents components, UriFormat format)
        {
            return _uri.GetComponents((GSystem.UriComponents)components, (GSystem.UriFormat)format);
        }

        public override int GetHashCode()
        { 
            return _uri.GetHashCode(); 
        }

        public string GetLeftPart(UriPartial part)
        {  
            return _uri.GetLeftPart((GSystem.UriPartial)part);
        }

        public new Type GetType()
        {
            return new Type(((GSystem.Object)this).GetType());
        }

        public bool IsBaseOf(Uri uri)
        {
            return _uri.IsBaseOf(uri.WrappedUri);
        }

        public bool IsWellFormedOriginalString()
        {
            return _uri.IsWellFormedOriginalString();
        }

        public Uri MakeRelativeUri(Uri uri)
        {
            return new Uri(_uri.MakeRelativeUri(uri.WrappedUri));
        }

        public override string ToString()
        { 
            return _uri.ToString(); 
        }

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            ((ISerializable)_uri).GetObjectData(info, context);
        }

        // Operators

        //
        //
        //  A static shortcut to Uri.Equals
        //
        [SecurityPermission(SecurityAction.InheritanceDemand, Flags = SecurityPermissionFlag.Infrastructure)]
        public static bool operator ==(Uri uri1, Uri uri2)
        {
            if ((object)uri1 == (object)uri2)
            {
                return true;
            }
            if ((object)uri1 == null || (object)uri2 == null)
            {
                return false;
            }
            return uri2.Equals(uri1);
        }

        //
        //
        //  A static shortcut to !Uri.Equals
        //
        [SecurityPermission(SecurityAction.InheritanceDemand, Flags = SecurityPermissionFlag.Infrastructure)]
        public static bool operator !=(Uri uri1, Uri uri2)
        {
            if ((object)uri1 == (object)uri2)
            {
                return false;
            }

            if ((object)uri1 == null || (object)uri2 == null)
            {
                return true;
            }

            return !uri2.Equals(uri1);
        }


    }
}
