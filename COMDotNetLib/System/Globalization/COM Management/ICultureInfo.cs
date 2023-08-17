// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo?view=netframework-4.8.1
// https://referencesource.microsoft.com/#mscorlib/system/globalization/cultureinfo.cs

using GGlobalization = global::System.Globalization;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;


namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("90801DAD-0FBF-47EA-82FD-15485DB4D971")]
    [Description("Provides information about a specific culture(called a locale for unmanaged code development). The information includes the names for the culture, the writing system, the calendar used, the sort order of strings, and formatting for dates and numbers.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICultureInfo
    {
        // Properties
        
        GGlobalization.CultureInfo CultureInfoObject
        { 
            get;
            set;
        }

        Calendar Calendar
        {
            [Description("Gets the default calendar used by the culture.")]
            get;
        }

        CompareInfo CompareInfo
        {
            [Description("Gets the CompareInfo that defines how to compare strings for the culture.")]
            get;
        }

        CultureTypes CultureTypes
        {
            [Description("Gets the culture types that pertain to the current CultureInfo object.")]
            get;
        }

        DateTimeFormatInfo DateTimeFormat
        {
            [Description("Gets or sets a DateTimeFormatInfo that defines the culturally appropriate format of displaying dates and times.")]
            get;
            [Description("Gets or sets a DateTimeFormatInfo that defines the culturally appropriate format of displaying dates and times.")]
            set;
        }

        string DisplayName
        {
            [Description("Gets the full localized culture name.")]
            get;
        }

        string EnglishName
        {
            [Description("Gets the culture name in the format languagefull [country/regionfull] in English.")]
            get;
        }

        bool IsNeutralCulture 
        {
            [Description("Gets a value indicating whether the current CultureInfo represents a neutral culture.")]
            get; 
        }

        bool IsReadOnly
        {
            [Description("Gets a value indicating whether the current CultureInfo is read-only.")]
            get;
        }

        int LCID
        {
            [Description("Gets the culture identifier for the current CultureInfo.")]
            get;
        }

        string Name
        {
            [Description("Gets the culture name in the format languagecode2-country/regioncode2.")]
            get;
        }

        string NativeName
        {
            [Description("Gets the culture name, consisting of the language, the country/region, and the optional script, that the culture is set to display.")]
            get;
        }

        NumberFormatInfo NumberFormat
        {
            [Description("Gets or sets a NumberFormatInfo that defines the culturally appropriate format of displaying numbers, currency, and percentage.")]
            get;
            [Description("Gets or sets a NumberFormatInfo that defines the culturally appropriate format of displaying numbers, currency, and percentage.")]
            set;
        }

        Calendar[] OptionalCalendars 
        {
            [Description("Gets the list of calendars that can be used by the culture.")]
            get; 
        }

        CultureInfo Parent
        {
            [Description("Gets the CultureInfo that represents the parent culture of the current CultureInfo.")]
            get;
        }

        TextInfo TextInfo
        {
            [Description("Gets the TextInfo that defines the writing system associated with the culture.")]
            get;
        }

        string ThreeLetterISOLanguageName
        {
            [Description("Gets the ISO 639-2 three-letter code for the language of the current CultureInfo.")]
            get;
        }

        string ThreeLetterWindowsLanguageName
        {
            [Description("Gets the three-letter code for the language as defined in the Windows API.")]
            get;
        }

        string TwoLetterISOLanguageName
        {
            [Description("Gets the ISO 639-1 two-letter or ISO 639-3 three-letter code for the language of the current CultureInfo.")]
            get;
        }

        bool UseUserOverride
        {
            [Description("Gets a value indicating whether the current CultureInfo object uses the user-selected culture settings.")]
            get;
        }

        //ICultureInfo CurrentUICulture { get; set; }

        // Methods

        [Description("Refreshes cached culture-related information.")]
        void ClearCachedData();

        [Description("Creates a copy of the current CultureInfo.")]
        object Clone ();

        [Description("Determines whether the specified object is the same culture as the current CultureInfo.")]
        bool Equals(object value);

        [Description("Gets an alternate user interface culture suitable for console applications when the default graphic user interface culture is unsuitable.")]
        CultureInfo GetConsoleFallbackUICulture();

        [Description("Gets the list of supported cultures filtered by the specified CultureTypes parameter.")]
        object GetFormat(Type formatType);

        [Description("Serves as a hash function for the current CultureInfo, suitable for hashing algorithms and data structures, such as a hash table.")]
        int GetHashCode();

        [Description("Returns a string containing the name of the current CultureInfo in the format languagecode2-country/regioncode2.")]
        string ToString();


        //https://referencesource.microsoft.com/#mscorlib/system/globalization/cultureinfo.cs,e319c6636909012f
        //public virtual Object GetFormat(Type formatType)
        //{
        //    if (formatType == typeof(NumberFormatInfo))
        //    {
        //        return (NumberFormat);
        //    }
        //    if (formatType == typeof(DateTimeFormatInfo))
        //    {
        //        return (DateTimeFormat);
        //    }
        //    return (null);
        //}

    }
}
