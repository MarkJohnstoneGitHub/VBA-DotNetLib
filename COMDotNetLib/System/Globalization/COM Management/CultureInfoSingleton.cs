using GGlobalization = global::System.Globalization;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("A4CD0781-B327-4B0A-90A5-F9944904282F")]
    [ProgId("DotNetLib.System.Globalization.CultureInfoSingleton")]
    [Description("CultureInfo factory methods and static members.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICultureInfoSingleton))]
    public class CultureInfoSingleton : ICultureInfoSingleton
    {
        public CultureInfoSingleton() { }

        // Factory Methods
        public CultureInfo Create(int culture)
        {
            return new CultureInfo(culture);
        }

        public CultureInfo Create2(string name)
        {
            return new CultureInfo(name);
        }

        public CultureInfo Create3(int culture, bool useUserOverride)
        {
            return new CultureInfo(culture, useUserOverride);
        }

        public CultureInfo Create4(string name, bool useUserOverride)
        {
            return new CultureInfo(name, useUserOverride);
        }

        // Properties
        public CultureInfo CurrentCulture
        {
            get
            {
                return CultureInfo.CurrentCulture;
            }
            set
            {
                CultureInfo.CurrentCulture = value;
            }
        }

        public CultureInfo CurrentUICulture
        {
            get
            {
                return CultureInfo.CurrentUICulture;
            }
            set
            {
                CultureInfo.CurrentUICulture = value;
            }
        }

        public CultureInfo DefaultThreadCurrentCulture
        {
            get
            {
                return CultureInfo.DefaultThreadCurrentUICulture;
            }
            set
            {
                CultureInfo.DefaultThreadCurrentCulture = value;
            }
        }

        public CultureInfo DefaultThreadCurrentUICulture
        {
            get
            {
                return CultureInfo.DefaultThreadCurrentUICulture;
            }
            set
            {
                CultureInfo.DefaultThreadCurrentUICulture = value;
            }
        }

        public CultureInfo InstalledUICulture
        {
            get
            {
                return CultureInfo.InstalledUICulture;
            }
        }

        public CultureInfo InvariantCulture
        {
            get
            {
                return CultureInfo.InvariantCulture;
            }
        }

        // Methods

        public CultureInfo CreateSpecificCulture(string name)
        {
            return CultureInfo.CreateSpecificCulture(name);
        }

        // GetCultureInfo Overloads
        public CultureInfo GetCultureInfo(int culture)
        {
            return CultureInfo.GetCultureInfo(culture);
        }

        public CultureInfo GetCultureInfo2(string name)
        {
            return CultureInfo.GetCultureInfo(name);
        }

        public CultureInfo GetCultureInfo3(string name, string altName)
        {
            return CultureInfo.GetCultureInfo(name, altName);
        }


        public CultureInfo GetCultureInfoByIetfLanguageTag(string name)
        {
            return CultureInfo.GetCultureInfoByIetfLanguageTag(name);
        }

        public CultureInfo[] GetCultures(GGlobalization.CultureTypes types)
        {
            return CultureInfo.GetCultures(types);
        }

        public CultureInfo ReadOnly(CultureInfo ci)
        {
            return CultureInfo.ReadOnly(ci);
        }

    }
}
