// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.regioninfo?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("8F1F39A8-C001-45AD-ABC6-F496FC4F34C4")]
    [Description("Contains information about the country/region.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRegionInfoSingleton
    {
        // Factory Methods

        [Description("Initializes a new instance of the RegionInfo class based on the country/region associated with the specified culture identifier.")]
        RegionInfo Create(int culture);

        [Description("Initializes a new instance of the RegionInfo class based on the country/region or specific culture, specified by name.")]
        RegionInfo Create2(string name);

        // Properties

        RegionInfo CurrentRegion 
        {
            [Description("Gets the RegionInfo that represents the country/region used by the current thread.")]
            get;
        }
    }
}
