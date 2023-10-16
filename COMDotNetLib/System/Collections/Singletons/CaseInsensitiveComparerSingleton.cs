using DotNetLib.System.Globalization;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Compares two objects for equivalence, ignoring the case of strings.")]
    [Guid("FF8EC7D9-9449-43B2-9A2D-8464EDD9E02F")]
    [ProgId("DotNetLib.System.Collections.CaseInsensitiveComparerSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICaseInsensitiveComparerSingleton))]
    public class CaseInsensitiveComparerSingleton : ICaseInsensitiveComparerSingleton
    {
        public CaseInsensitiveComparerSingleton() { }

        public CaseInsensitiveComparer Create()
        {
            return new CaseInsensitiveComparer();
        }

        public CaseInsensitiveComparer Create2(CultureInfo culture)
        {
            return new CaseInsensitiveComparer(culture.WrappedCultureInfo);
        }

        public CaseInsensitiveComparer Default => CaseInsensitiveComparer.Default;

        public CaseInsensitiveComparer DefaultInvariant => CaseInsensitiveComparer.DefaultInvariant;


    }
}
