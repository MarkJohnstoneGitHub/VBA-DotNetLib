// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.regioninfo?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("15BBCCEF-698F-4DE2-9C7F-139E772C0FA3")]
    [Description("Contains information about the country/region.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRegionInfo
    {
        string CurrencyEnglishName 
        {
            [Description("Gets the name, in English, of the currency used in the country/region.")]
            get;
        }

        string CurrencyNativeName 
        {
            [Description("Gets the name of the currency used in the country/region, formatted in the native language of the country/region.")]
            get;
        }

        string CurrencySymbol 
        {
            [Description("Gets the currency symbol associated with the country/region.")]
            get;
        }

        string DisplayName 
        {
            [Description("Gets the full name of the country/region in the language of the localized version of .NET.")]
            get;
        }

        string EnglishName 
        {
            [Description("Gets the full name of the country/region in English.")]
            get;
        }

        int GeoId 
        {
            [Description("Gets a unique identification number for a geographical region, country, city, or location.")]
            get;
        }

        bool IsMetric 
        {
            [Description("Gets a value indicating whether the country/region uses the metric system for measurements.")]
            get;
        }

        string ISOCurrencySymbol 
        {
            [Description("Gets the three-character ISO 4217 currency symbol associated with the country/region.")]
            get;
        }

        string Name 
        {
            [Description("Gets the name or ISO 3166 two-letter country/region code for the current RegionInfo object.")]
            get;
        }

        string NativeName 
        {
            [Description("Gets the name of a country/region formatted in the native language of the country/region.")]
            get;
        }

        string ThreeLetterISORegionName 
        {
            [Description("Gets the three-letter code defined in ISO 3166 for the country/region.")]
            get;
        }

        string ThreeLetterWindowsRegionName 
        {
            [Description("Gets the three-letter code assigned by Windows to the country/region represented by this RegionInfo.")]
            get;
        }

        string TwoLetterISORegionName 
        {
            [Description("Gets the two-letter code defined in ISO 3166 for the country/region.")]
            get;
        }

        // Methods

        [Description("Determines whether the specified object is the same instance as the current RegionInfo.")]
        bool Equals(object value);

        [Description("Serves as a hash function for the current RegionInfo, suitable for hashing algorithms and data structures, such as a hash table.")]
        int GetHashCode();

        [Description("Returns a string containing the culture name or ISO 3166 two-letter country/region codes specified for the current RegionInfo.")] 
        string ToString();

    }
}
