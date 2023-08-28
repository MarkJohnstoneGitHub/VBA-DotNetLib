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
    public class CultureInfo : ICloneable, IFormatProvider, ICultureInfo
    {
        private GGlobalization.CultureInfo _cultureInfo;
        private NumberFormatInfo _numberFormatInfo;

        //private static ICultureInfo _currentCulture = new CultureInfo(GGlobalization.CultureInfo.CurrentCulture);
        //private static ICultureInfo _currentUICulture = new CultureInfo(GGlobalization.CultureInfo.CurrentUICulture);
        //private static ICultureInfo _defaultThreadCurrentCulture = new CultureInfo(GGlobalization.CultureInfo.DefaultThreadCurrentCulture);
        //private static ICultureInfo _defaultThreadCurrentUICulture = new CultureInfo(GGlobalization.CultureInfo.DefaultThreadCurrentUICulture);
        //private static ICultureInfo _installedUICulture = new CultureInfo(GGlobalization.CultureInfo.InstalledUICulture);
        //private static ICultureInfo _invariantCulture = new CultureInfo(GGlobalization.CultureInfo.InvariantCulture);

        //Constructors

        internal CultureInfo(GGlobalization.CultureInfo cultureInfo)
        {
            this._cultureInfo = cultureInfo;
            _numberFormatInfo = new NumberFormatInfo(_cultureInfo.NumberFormat);
        }

        //public CultureInfo()
        //{
        //}

        public CultureInfo(int culture)
        {
            this._cultureInfo = new GGlobalization.CultureInfo(culture);
            _numberFormatInfo = new NumberFormatInfo(_cultureInfo.NumberFormat);
        }

        public CultureInfo(string name)
        {
            this._cultureInfo = new GGlobalization.CultureInfo(name);
            _numberFormatInfo = new NumberFormatInfo(_cultureInfo.NumberFormat);
        }

        public CultureInfo(int culture, bool useUserOverride)
        {
            this._cultureInfo = new GGlobalization.CultureInfo(culture, useUserOverride);
            _numberFormatInfo = new NumberFormatInfo(_cultureInfo.NumberFormat);
        }

        public CultureInfo(string name, bool useUserOverride)
        {
            this._cultureInfo = new GGlobalization.CultureInfo(name, useUserOverride);
            _numberFormatInfo = new NumberFormatInfo(_cultureInfo.NumberFormat);
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
            get => _numberFormatInfo;
            set => _numberFormatInfo = value;
        }

        public Calendar[] OptionalCalendars => _cultureInfo.OptionalCalendars;

        public CultureInfo Parent
        {
            get { return new CultureInfo(_cultureInfo.Parent); }
        }

        public TextInfo TextInfo => _cultureInfo.TextInfo;

        public string ThreeLetterISOLanguageName => _cultureInfo.ThreeLetterISOLanguageName;

        public string ThreeLetterWindowsLanguageName => _cultureInfo.ThreeLetterWindowsLanguageName;

        public string TwoLetterISOLanguageName => _cultureInfo.TwoLetterISOLanguageName;

        public bool UseUserOverride => _cultureInfo.UseUserOverride;

        public static CultureInfo CurrentCulture
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

        public static CultureInfo CurrentUICulture
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

        public static CultureInfo DefaultThreadCurrentCulture
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

        public static CultureInfo DefaultThreadCurrentUICulture
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
        public static CultureInfo InstalledUICulture
        {
            get
            {
                return new CultureInfo(GGlobalization.CultureInfo.InstalledUICulture);
            }
        }

        // Static
        public static CultureInfo InvariantCulture
        {
            get
            {
                return new CultureInfo(GGlobalization.CultureInfo.InvariantCulture);
            }
        }

        // Methods

        public CultureInfo GetConsoleFallbackUICulture()
        {
            return new CultureInfo(_cultureInfo.GetConsoleFallbackUICulture());
        }

        // https://stackoverflow.com/questions/24413077/what-is-the-best-way-to-compare-two-cultureinfo-instances
        public override bool Equals(object value)
        {
            return value is CultureInfo ci && this.CultureInfoObject == ci.CultureInfoObject;
        }

        public override int GetHashCode() 
        { 
            return this.GetHashCode(); 
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
        public static CultureInfo CreateSpecificCulture(string name)
        {
            return new CultureInfo(GGlobalization.CultureInfo.CreateSpecificCulture(name));
        }

        public static CultureInfo GetCultureInfo(int culture)
        {
            return new CultureInfo(GGlobalization.CultureInfo.GetCultureInfo(culture));
        }

        public static CultureInfo GetCultureInfo(string name)
        {
            return new CultureInfo(GGlobalization.CultureInfo.GetCultureInfo(name));
        }

        public static CultureInfo GetCultureInfo(string name, string altName)
        {
            return new CultureInfo(GGlobalization.CultureInfo.GetCultureInfo(name, altName));
        }

        public static CultureInfo GetCultureInfoByIetfLanguageTag(string name)
        {
            return new CultureInfo(GGlobalization.CultureInfo.GetCultureInfoByIetfLanguageTag(name));
        }

        public static CultureInfo[] GetCultures(GGlobalization.CultureTypes types)
        {
            GGlobalization.CultureInfo[] cultures = GGlobalization.CultureInfo.GetCultures(types);
            //https://stackoverflow.com/questions/9917390/how-to-create-and-initialize-an-array-with-another-array
            CultureInfo[] output = new CultureInfo[cultures.Length];

            int index = 0;
            foreach (GGlobalization.CultureInfo culture in cultures)
            {
                output[index] = new CultureInfo(culture);
                index++;
            }
            return output;
        }

        // Static
        public static CultureInfo ReadOnly(CultureInfo ci)
        {
            return new CultureInfo(GGlobalization.CultureInfo.ReadOnly(ci.CultureInfoObject));
        }

    }
}
