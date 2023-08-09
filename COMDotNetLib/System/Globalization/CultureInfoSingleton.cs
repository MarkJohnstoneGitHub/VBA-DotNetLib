using GGlobalization = global::System.Globalization;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DotNetLib.System.Globalization
{
    [ComVisible(false)]
    [Guid("A4CD0781-B327-4B0A-90A5-F9944904282F")]
    [ProgId("DotNetLib.System.Globalization.CultureInfoSingleton")]
    [Description("Provides information about a specific culture(called a locale for unmanaged code development). The information includes the names for the culture, the writing system, the calendar used, the sort order of strings, and formatting for dates and numbers.")]
    [ClassInterface(ClassInterfaceType.None)]
    public class CultureInfoSingleton //: ICultureInfoSingleton
    {
        private static ICultureInfo _currentCulture = new CultureInfo(GGlobalization.CultureInfo.CurrentCulture);
        private static ICultureInfo _currentUICulture = new CultureInfo(GGlobalization.CultureInfo.CurrentUICulture);
        private static ICultureInfo _defaultThreadCurrentCulture = new CultureInfo(GGlobalization.CultureInfo.DefaultThreadCurrentCulture);
        private static ICultureInfo _defaultThreadCurrentUICulture = new CultureInfo(GGlobalization.CultureInfo.DefaultThreadCurrentUICulture);
        private static ICultureInfo _installedUICulture = new CultureInfo(GGlobalization.CultureInfo.InstalledUICulture);
        private static ICultureInfo _invariantCulture = new CultureInfo(GGlobalization.CultureInfo.InvariantCulture);

        public CultureInfoSingleton() { }

        // Factory Methods
        public ICultureInfo Create(int culture)
        {
            return new CultureInfo(culture);
        }

        public ICultureInfo Create2(string name)
        {
            return new CultureInfo(name);
        }

        public ICultureInfo Create3(int culture, bool useUserOverride)
        {
            return new CultureInfo(culture, useUserOverride);
        }

        public ICultureInfo Create4(string name, bool useUserOverride)
        {
            return new CultureInfo(name, useUserOverride);
        }

        // Static
        //public ICultureInfo CurrentCulture
        //{
        //    get
        //    {
        //        return (CultureInfo)GGlobalization.CultureInfo.CurrentCulture;
        //    }
        //    set
        //    {
        //        GGlobalization.CultureInfo.CurrentCulture = (GGlobalization.CultureInfo)value;
        //    }
        //}

        //public ICultureInfo CurrentUICulture
        //{
        //    get => (CultureInfo)GGlobalization.CultureInfo.CurrentUICulture;
        //    set
        //    {
        //        GGlobalization.CultureInfo.CurrentUICulture = (GGlobalization.CultureInfo)value;
        //    }
        //}

        //// Static
        //public ICultureInfo DefaultThreadCurrentCulture
        //{
        //    get
        //    {
        //        return (CultureInfo)GGlobalization.CultureInfo.DefaultThreadCurrentCulture;
        //    }
        //    set
        //    {
        //        GGlobalization.CultureInfo.DefaultThreadCurrentCulture = (GGlobalization.CultureInfo)value;
        //    }
        //}

        //// Static
        //public ICultureInfo DefaultThreadCurrentUICulture
        //{
        //    get
        //    {
        //        return (CultureInfo)GGlobalization.CultureInfo.DefaultThreadCurrentUICulture;
        //    }
        //    set
        //    {
        //        GGlobalization.CultureInfo.DefaultThreadCurrentUICulture = (GGlobalization.CultureInfo)value;
        //    }
        //}

        //// Static
        //public ICultureInfo InstalledUICulture
        //{
        //    get
        //    {
        //        return (CultureInfo)GGlobalization.CultureInfo.InstalledUICulture;
        //    }
        //}

        //// Static
        //public ICultureInfo InvariantCulture
        //{
        //    get
        //    {
        //        return (CultureInfo)GGlobalization.CultureInfo.InvariantCulture;
        //    }
        //}

        //// Methods

        //// Static
        //public ICultureInfo CreateSpecificCulture(string name)
        //{
        //    return (CultureInfo)GGlobalization.CultureInfo.CreateSpecificCulture(name);
        //}

        //// GetCultureInfo Overloads
        //// Static
        //public ICultureInfo GetCultureInfo(int culture)
        //{
        //    return (CultureInfo)GGlobalization.CultureInfo.GetCultureInfo(culture);
        //}
        //// Static
        //public ICultureInfo GetCultureInfo2(string name)
        //{
        //    return (CultureInfo)GGlobalization.CultureInfo.GetCultureInfo(name);
        //}
        //// Static
        //public ICultureInfo GetCultureInfo3(string name, string altName)
        //{
        //    return (CultureInfo)GGlobalization.CultureInfo.GetCultureInfo(name, altName);
        //}

        //// Static
        //public ICultureInfo GetCultureInfoByIetfLanguageTag(string name)
        //{

        //    return new CultureInfo(GGlobalization.CultureInfo.InvariantCulture.LCID)
        //    //return (CultureInfo)GGlobalization.CultureInfo.GetCultureInfoByIetfLanguageTag(name);
        //}

        //// Static
        //public ICultureInfo[] GetCultures(GGlobalization.CultureTypes types)
        //{
        //    return (CultureInfo[])GGlobalization.CultureInfo.GetCultures(types);
        //}

        //// Static
        //public ICultureInfo ReadOnly(ICultureInfo ci)
        //{
        //    return (CultureInfo)GGlobalization.CultureInfo.ReadOnly((GGlobalization.CultureInfo)ci);
        //}

    }
}
