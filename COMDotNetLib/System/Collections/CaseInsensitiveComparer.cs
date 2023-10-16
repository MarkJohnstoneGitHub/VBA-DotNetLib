// https://learn.microsoft.com/en-us/dotnet/api/system.collections.caseinsensitivecomparer.compare?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System.Globalization;
using DotNetLib.Extensions;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Compares two objects for equivalence, ignoring the case of strings.")]
    [Guid("66AD88CF-B4BD-45DA-B01B-FE96BCDED695")]
    [ProgId("DotNetLib.System.Collections.CaseInsensitiveComparer")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICaseInsensitiveComparer))]
    public class CaseInsensitiveComparer : ICaseInsensitiveComparer, GCollections.IComparer
    {
        private GCollections.CaseInsensitiveComparer _caseInsensitiveComparer;


        public CaseInsensitiveComparer() 
        {
            _caseInsensitiveComparer = new GCollections.CaseInsensitiveComparer();
        }

        public CaseInsensitiveComparer(CultureInfo culture)
        {
            _caseInsensitiveComparer = new GCollections.CaseInsensitiveComparer(culture);
        }

        public CaseInsensitiveComparer(GCollections.CaseInsensitiveComparer caseInsensitiveComparer)
        {
            _caseInsensitiveComparer = caseInsensitiveComparer;
        }

        public static CaseInsensitiveComparer Default => new CaseInsensitiveComparer(GCollections.CaseInsensitiveComparer.Default);

        public static CaseInsensitiveComparer DefaultInvariant => new CaseInsensitiveComparer(GCollections.CaseInsensitiveComparer.DefaultInvariant);

        public int Compare(object x, object y)
        {
            return _caseInsensitiveComparer.Compare(x, y);
        }
    }
}
