// https://learn.microsoft.com/en-us/dotnet/api/system.stringcomparer?view=netframework-4.8.1

using DotNetLib.System.Globalization;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("135BBD2F-2DF5-4070-A7D5-3A4DFE788CBF")]
    [Description("Represents a string comparison operation that uses specific case and culture-based or ordinal comparison rules.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStringComparerSingleton
    {
        StringComparer CurrentCulture 
        {
            [Description("Gets a StringComparer object that performs a case-sensitive string comparison using the word comparison rules of the current culture.")]
            get;
        }

        StringComparer CurrentCultureIgnoreCase 
        {
            [Description("Gets a StringComparer object that performs case-insensitive string comparisons using the word comparison rules of the current culture.")]
            get;
        }

        StringComparer InvariantCulture 
        {
            [Description("Gets a StringComparer object that performs a case-sensitive string comparison using the word comparison rules of the invariant culture.")]
            get;
        }


        StringComparer InvariantCultureIgnoreCase 
        {
            [Description("Gets a StringComparer object that performs a case-insensitive string comparison using the word comparison rules of the invariant culture.")]
            get;
        }


        StringComparer Ordinal 
        {
            [Description("Gets a StringComparer object that performs a case-sensitive ordinal string comparison.")]
            get;
        }

        StringComparer OrdinalIgnoreCase 
        {

            [Description("Gets a StringComparer object that performs a case-insensitive ordinal string comparison.")]
            get;
        }

        // Methods

        [Description("Creates a StringComparer object that compares strings according to the rules of a specified culture.")]
        StringComparer Create(CultureInfo culture, bool ignoreCase);

    }
}
