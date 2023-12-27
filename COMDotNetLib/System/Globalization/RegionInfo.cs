// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.regioninfo?view=netframework-4.8.1

using GGlobalization = global::System.Globalization;
using System.Threading.Tasks;
using DotNetLib.Extensions;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("0A621151-DEA7-4091-96A0-B021621E6C2B")]
    [ProgId("DotNetLib.System.Globalization.RegionInfo")]
    [Description("Contains information about the country/region.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IRegionInfo))]
    public class RegionInfo : IRegionInfo, IWrappedObject
    {
        private GGlobalization.RegionInfo _regionInfo;

        public RegionInfo(int culture)
        {
            _regionInfo = new GGlobalization.RegionInfo(culture);
        }

        internal RegionInfo(GGlobalization.RegionInfo regionInfo)
        {
            _regionInfo = regionInfo;
        }

        public RegionInfo(string name)
        {
            _regionInfo =  new GGlobalization.RegionInfo(name);
        }

        public GGlobalization.RegionInfo WrappedRegionInfo => _regionInfo;

        public object WrappedObject => _regionInfo;

        public virtual string CurrencyEnglishName => _regionInfo.CurrencyEnglishName;

        public virtual string CurrencyNativeName => _regionInfo.CurrencyNativeName;

        public virtual string CurrencySymbol => _regionInfo.CurrencySymbol;

        public static RegionInfo CurrentRegion => new RegionInfo(GGlobalization.RegionInfo.CurrentRegion);

        public virtual string DisplayName =>  _regionInfo.DisplayName;

        public virtual string EnglishName => _regionInfo.EnglishName;

        public virtual int GeoId => _regionInfo.GeoId;

        public virtual bool IsMetric => _regionInfo.IsMetric;

        public virtual string ISOCurrencySymbol => _regionInfo.ISOCurrencySymbol;

        public virtual string Name => _regionInfo.Name;

        public virtual string NativeName => _regionInfo.NativeName;

        public virtual string ThreeLetterISORegionName => _regionInfo.ThreeLetterISORegionName;

        public virtual string ThreeLetterWindowsRegionName => _regionInfo.ThreeLetterWindowsRegionName;

        public virtual string TwoLetterISORegionName => _regionInfo?.TwoLetterISORegionName;

        public override bool Equals(object value)
        { 
            return _regionInfo.Equals(value.Unwrap()); 
        }

        public override int GetHashCode()
        { 
            return _regionInfo.GetHashCode(); 
        }

        public override string ToString()
        {
            return _regionInfo.ToString();
        }
    }


}
