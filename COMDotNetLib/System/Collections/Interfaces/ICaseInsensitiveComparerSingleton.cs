// https://learn.microsoft.com/en-us/dotnet/api/system.collections.caseinsensitivecomparer.compare?view=netframework-4.8.1

using DotNetLib.System.Globalization;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("73E4DB6A-04C0-419B-8A2B-E97CA14B9FC0")]
    [Description("Compares two objects for equivalence, ignoring the case of strings.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICaseInsensitiveComparerSingleton
    {
        // Factorty Methods
        [Description("Initializes a new instance of the CaseInsensitiveComparer class using the CurrentCulture of the current thread.")]
        CaseInsensitiveComparer Create();

        [Description("Initializes a new instance of the CaseInsensitiveComparer class using the specified CultureInfo.")]
        CaseInsensitiveComparer Create2(CultureInfo culture);

        // Properties

        
        CaseInsensitiveComparer Default 
        {
            [Description("Gets an instance of CaseInsensitiveComparer that is associated with the CurrentCulture of the current thread and that is always available.")]
            get;
        }

        
        CaseInsensitiveComparer DefaultInvariant
        {
            [Description("Gets an instance of CaseInsensitiveComparer that is associated with InvariantCulture and that is always available.")]
            get;
        }

    }
}
