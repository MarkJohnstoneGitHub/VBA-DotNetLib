// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo?view=netframework-4.8.1
// https://referencesource.microsoft.com/#mscorlib/system/globalization/cultureinfo.cs

using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using GGlobalization = global::System.Globalization;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("2958A3AC-4018-4BAE-ACB6-420F6BC9CB60")]
    [ProgId("DotNetLib.System.Globalization.CultureInfo")]
    [Description("Provides information about a specific culture(called a locale for unmanaged code development). The information includes the names for the culture, the writing system, the calendar used, the sort order of strings, and formatting for dates and numbers.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICultureInfo))]
    public class CultureInfo : ICultureInfo, ICultureInfoSingleton
    {
        private GGlobalization.CultureInfo _cultureInfo;

        //Constructors

        internal CultureInfo(GGlobalization.CultureInfo cultureInfo)
        {
            this._cultureInfo = cultureInfo;
        }

        public CultureInfo()
        {
        }

        public CultureInfo(int culture)
        {
            this._cultureInfo = new GGlobalization.CultureInfo(culture);
        }

        public CultureInfo(string name)
        {
            this._cultureInfo = new GGlobalization.CultureInfo(name);
        }

        public CultureInfo(int culture, bool useUserOverride)
        {
            this._cultureInfo = new GGlobalization.CultureInfo(culture, useUserOverride);
        }

        public CultureInfo(string name, bool useUserOverride)
        {
            this._cultureInfo = new GGlobalization.CultureInfo(name, useUserOverride);
        }


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


        // Properties

        public GGlobalization.CultureInfo CultureInfoObject
        {
            get { return this._cultureInfo; }
            set { this._cultureInfo = value; }
        }

        public Calendar Calendar
        {
            get { return _cultureInfo.Calendar; }
        }

        public CompareInfo CompareInfo
        {
            get { return _cultureInfo.CompareInfo; }
        }

        public CultureTypes CultureTypes
        {
            get { return _cultureInfo.CultureTypes; }
        }

        public DateTimeFormatInfo DateTimeFormat 
        { 
            get => _cultureInfo.DateTimeFormat; 
            set => _cultureInfo.DateTimeFormat = value; 
        }

        public string DisplayName => _cultureInfo.DisplayName;

        public string EnglishName => _cultureInfo.EnglishName;

        public bool IsNeutralCulture => _cultureInfo.IsNeutralCulture;

        public bool IsReadOnly => _cultureInfo.IsReadOnly;

        public int LCID => _cultureInfo.LCID;

        public string Name => _cultureInfo.Name;

        public string NativeName => _cultureInfo.NativeName;

        public NumberFormatInfo NumberFormat 
        { 
            get => _cultureInfo.NumberFormat;
            set => _cultureInfo.NumberFormat = value;
        }

        public Calendar[] OptionalCalendars => _cultureInfo.OptionalCalendars;

        public ICultureInfo Parent
        {
            get { return new CultureInfo(_cultureInfo.Parent); }
        }

        public TextInfo TextInfo => _cultureInfo.TextInfo;

        public string ThreeLetterISOLanguageName => _cultureInfo.ThreeLetterISOLanguageName;

        public string ThreeLetterWindowsLanguageName => _cultureInfo.ThreeLetterWindowsLanguageName;

        public string TwoLetterISOLanguageName => _cultureInfo.TwoLetterISOLanguageName;

        public bool UseUserOverride => _cultureInfo.UseUserOverride;

        // Static
        public ICultureInfo CurrentCulture
        {
            get
            {
                return new CultureInfo(GGlobalization.CultureInfo.CurrentCulture);
            }
            set
            {
                GGlobalization.CultureInfo.CurrentCulture = value.CultureInfoObject;
            }
        }

        // Static
        public ICultureInfo CurrentUICulture
        {
            get
            {
                return new CultureInfo(GGlobalization.CultureInfo.CurrentUICulture);
            }
            set
            {
                GGlobalization.CultureInfo.CurrentUICulture = value.CultureInfoObject;
            }
        }

        // Static
        public ICultureInfo DefaultThreadCurrentCulture
        {
            get
            {
                return new CultureInfo(GGlobalization.CultureInfo.DefaultThreadCurrentUICulture);
            }
            set
            {
                GGlobalization.CultureInfo.DefaultThreadCurrentCulture = value.CultureInfoObject;
            }
        }

        // Static
        public ICultureInfo DefaultThreadCurrentUICulture
        {
            get
            {
                return new CultureInfo(GGlobalization.CultureInfo.DefaultThreadCurrentUICulture);
            }
            set
            {
                GGlobalization.CultureInfo.DefaultThreadCurrentUICulture = value.CultureInfoObject;
            }
        }

        // Static
        public ICultureInfo InstalledUICulture
        {
            get
            {
                return new CultureInfo(GGlobalization.CultureInfo.InstalledUICulture);
            }
        }

        // Static
        public ICultureInfo InvariantCulture
        {
            get
            {
                return new CultureInfo(GGlobalization.CultureInfo.InvariantCulture);
            }
        }


        // Methods

        public ICultureInfo GetConsoleFallbackUICulture()
        {
            return new CultureInfo(_cultureInfo.GetConsoleFallbackUICulture());
        }

        public override string ToString()
        {
            return _cultureInfo.ToString();
        }

        public void ClearCachedData()
        {
            _cultureInfo.ClearCachedData();
        }

        public object Clone()
        {
            return  (CultureInfo)_cultureInfo.Clone();

        }

        public object GetFormat(Type formatType)
        {
            return _cultureInfo.GetFormat(formatType);
        }


        // Static
        public ICultureInfo CreateSpecificCulture(string name)
        {
            return new CultureInfo(GGlobalization.CultureInfo.CreateSpecificCulture(name));
        }

        // GetCultureInfo Overloads
        // Static
        public ICultureInfo GetCultureInfo(int culture)
        {
            return new CultureInfo(GGlobalization.CultureInfo.GetCultureInfo(culture)); 
        }
        // Static
        public ICultureInfo GetCultureInfo2(string name)
        {
            return new CultureInfo(GGlobalization.CultureInfo.GetCultureInfo(name));
        }
        // Static
        public ICultureInfo GetCultureInfo3(string name, string altName)
        {
            return new CultureInfo(GGlobalization.CultureInfo.GetCultureInfo(name, altName)); 
        }

        // Static
        public ICultureInfo GetCultureInfoByIetfLanguageTag(string name)
        {
            return new CultureInfo(GGlobalization.CultureInfo.GetCultureInfoByIetfLanguageTag(name)); 
        }

        // Static
        public ICultureInfo[] GetCultures(GGlobalization.CultureTypes types)
        {
            GGlobalization.CultureInfo[] cultures = GGlobalization.CultureInfo.GetCultures(types);
            //https://stackoverflow.com/questions/9917390/how-to-create-and-initialize-an-array-with-another-array
            ICultureInfo[] output = new ICultureInfo[cultures.Length];

            int index = 0;
            foreach (GGlobalization.CultureInfo culture in cultures)
            {
                output[index] = new CultureInfo(culture);
                index++;
            }
            return output;
        }

        // Static
        public ICultureInfo ReadOnly(ICultureInfo ci)
        {
            return new CultureInfo(GGlobalization.CultureInfo.ReadOnly((GGlobalization.CultureInfo)ci));
        }



    }
}
