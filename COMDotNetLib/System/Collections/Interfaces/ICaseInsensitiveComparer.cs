// https://learn.microsoft.com/en-us/dotnet/api/system.collections.caseinsensitivecomparer.compare?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("D24C309E-74E4-430A-BE43-59F1E1499815")]
    [Description("Compares two objects for equivalence, ignoring the case of strings.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICaseInsensitiveComparer
    {
        [Description("Performs a case-insensitive comparison of two objects of the same type and returns a value indicating whether one is less than, equal to, or greater than the other.")]
        int Compare(object x, object y);
    }
}
