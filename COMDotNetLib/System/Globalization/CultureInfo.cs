// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo?view=netframework-4.8.1
// https://referencesource.microsoft.com/#mscorlib/system/globalization/cultureinfo.cs

using DotNetLib.Extensions;
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
        private DateTimeFormatInfo _dateTimeFormatInfo;
        private TextInfo _textInfo;
        private ICalendar _calendar;
        private ICalendar[] _optionalCalendars;

        //private static ICultureInfo _currentCulture = new CultureInfo(GGlobalization.CultureInfo.CurrentCulture);
        //private static ICultureInfo _currentUICulture = new CultureInfo(GGlobalization.CultureInfo.CurrentUICulture);
        //private static ICultureInfo _defaultThreadCurrentCulture = new CultureInfo(GGlobalization.CultureInfo.DefaultThreadCurrentCulture);
        //private static ICultureInfo _defaultThreadCurrentUICulture = new CultureInfo(GGlobalization.CultureInfo.DefaultThreadCurrentUICulture);
        //private static ICultureInfo _installedUICulture = new CultureInfo(GGlobalization.CultureInfo.InstalledUICulture);
        //private static ICultureInfo _invariantCulture = new CultureInfo(GGlobalization.CultureInfo.InvariantCulture);

        //Constructors

        public CultureInfo(GGlobalization.CultureInfo cultureInfo)
        {
            WrappedCultureInfo = cultureInfo;
        }

        //public CultureInfo()
        //{
        //}

        public CultureInfo(int culture)
        {
            WrappedCultureInfo = new GGlobalization.CultureInfo(culture);
        }

        public CultureInfo(string name)
        {
            WrappedCultureInfo = new GGlobalization.CultureInfo(name);
        }

        public CultureInfo(int culture, bool useUserOverride)
        {
            WrappedCultureInfo = new GGlobalization.CultureInfo(culture, useUserOverride);
        }

        public CultureInfo(string name, bool useUserOverride)
        {
            WrappedCultureInfo = new GGlobalization.CultureInfo(name, useUserOverride);
        }

        // Properties

        public GGlobalization.CultureInfo WrappedCultureInfo
        {
            get => _cultureInfo;
            set 
            { 
                _cultureInfo = value;
                _numberFormatInfo = new NumberFormatInfo(_cultureInfo.NumberFormat);
                _dateTimeFormatInfo = new DateTimeFormatInfo(_cultureInfo.DateTimeFormat);
                _textInfo = new TextInfo(_cultureInfo.TextInfo);
                _calendar = _cultureInfo.Calendar.Wrap();
                _optionalCalendars = _cultureInfo.OptionalCalendars.Wrap();
            }
        }

        //public Calendar Calendar => _cultureInfo.Calendar;

        public ICalendar Calendar => _calendar;

        public CompareInfo CompareInfo => _cultureInfo.CompareInfo;

        public CultureTypes CultureTypes  => _cultureInfo.CultureTypes;

        public DateTimeFormatInfo DateTimeFormat
        {
            get => _dateTimeFormatInfo;
            set 
            {
                _cultureInfo.DateTimeFormat = value.WrappedDateTimeFormatInfo;
                _dateTimeFormatInfo = value;
            }
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
            set
            {
                _numberFormatInfo = value;
                _cultureInfo.NumberFormat = value.WrappedNumberFormatInfo;
            }
        }

        //public Calendar[] OptionalCalendars => _cultureInfo.OptionalCalendars;

        public ICalendar[] OptionalCalendars => _optionalCalendars;


        //TODO : Check implementation return new?
        public CultureInfo Parent => new CultureInfo(_cultureInfo.Parent);

        public TextInfo TextInfo => _textInfo; 

        public string ThreeLetterISOLanguageName => _cultureInfo.ThreeLetterISOLanguageName;

        public string ThreeLetterWindowsLanguageName => _cultureInfo.ThreeLetterWindowsLanguageName;

        public string TwoLetterISOLanguageName => _cultureInfo.TwoLetterISOLanguageName;

        public bool UseUserOverride => _cultureInfo.UseUserOverride;

        public static CultureInfo CurrentCulture
        {
            get => new CultureInfo(GGlobalization.CultureInfo.CurrentCulture); 
            set { GGlobalization.CultureInfo.CurrentCulture = value.WrappedCultureInfo; }
        }

        public static CultureInfo CurrentUICulture
        {
            get => new CultureInfo(GGlobalization.CultureInfo.CurrentUICulture);
            set { GGlobalization.CultureInfo.CurrentUICulture = value.WrappedCultureInfo; }
        }

        public static CultureInfo DefaultThreadCurrentCulture
        {
            get => new CultureInfo(GGlobalization.CultureInfo.DefaultThreadCurrentUICulture);
            set { GGlobalization.CultureInfo.DefaultThreadCurrentCulture = value.WrappedCultureInfo;}
        }

        public static CultureInfo DefaultThreadCurrentUICulture
        {
            get => new CultureInfo(GGlobalization.CultureInfo.DefaultThreadCurrentUICulture);
            set { GGlobalization.CultureInfo.DefaultThreadCurrentUICulture = value.WrappedCultureInfo; }
        }

        public static CultureInfo InstalledUICulture => new CultureInfo(GGlobalization.CultureInfo.InstalledUICulture);

        public static CultureInfo InvariantCulture  => new CultureInfo(GGlobalization.CultureInfo.InvariantCulture);


        // Methods

        public CultureInfo GetConsoleFallbackUICulture()
        {
            return new CultureInfo(_cultureInfo.GetConsoleFallbackUICulture());
        }

        // https://stackoverflow.com/questions/24413077/what-is-the-best-way-to-compare-two-cultureinfo-instances
        public override bool Equals(object value)
        {
            return value is CultureInfo ci && WrappedCultureInfo == ci.WrappedCultureInfo;
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

        //TODO : Check implementation
        // Also check if type is typeof mscorlib.NumberFormatInfo, mscorlib.DateTimeFormatInfo ?
        public object GetFormat(Type formatType)
        {
            if (formatType == typeof(NumberFormatInfo))
            {
                return NumberFormat;
                //return _cultureInfo.NumberFormat;   
            }
            if (formatType == typeof(DateTimeFormatInfo))
            {
                return DateTimeFormat;
                //return _cultureInfo.DateTimeFormat; 
            }
            return null;
        }

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
            return new CultureInfo(GGlobalization.CultureInfo.ReadOnly(ci.WrappedCultureInfo));
        }

    }
}
