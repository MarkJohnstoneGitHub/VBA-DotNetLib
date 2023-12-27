// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.regioninfo?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("6C499B1D-D1E0-4FB3-8738-DA13B5C63C55")]
    [ProgId("DotNetLib.System.Globalization.RegionInfoSingleton")]
    [Description("Contains information about the country/region.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRegionInfoSingleton))]
    public class RegionInfoSingleton : IRegionInfoSingleton
    {
        public RegionInfoSingleton() { }

        // Factory  Methods
        public RegionInfo Create(int culture)
        {
            return  new RegionInfo(culture);
        }

        public RegionInfo Create2(string name)
        {
            return new RegionInfo(name);
        }

        // Properties

        public RegionInfo CurrentRegion => RegionInfo.CurrentRegion;


    }
}
