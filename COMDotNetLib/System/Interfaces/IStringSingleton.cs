// https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using CultureInfo = DotNetLib.System.Globalization.CultureInfo;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("4F39D64D-9BD3-4AF7-A124-7A88364BE29F")]
    [Description("")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStringSingleton
    {
        // Fields

        string EmptyString
        {
            [Description("Represents the empty string. This field is read-only.")]
            get;
        }

        // Compare Overloads

        [Description("Compares two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order.")]
        int Compare(string strA, string strB, bool ignoreCase = false);

        [Description("Compares two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order.")]
        int Compare2(string strA, string strB, StringComparison comparisonType);

        [Description("Compares two specified String objects, ignoring or honoring their case, and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order.")]
        int Compare3(string strA, string strB, bool ignoreCase, CultureInfo culture);

        [Description("Compares two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two strings to each other in the sort order.")]
        int Compare4(string strA, string strB, CultureInfo culture, CompareOptions options);

        [Description("Compares substrings of two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order.")]
        int Compare5(string strA, int indexA, string strB, int indexB, int length, bool ignoreCase = false);

        [Description("Compares substrings of two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order.")]
        int Compare6(string strA, int indexA, string strB, int indexB, int length, StringComparison comparisonType);

        [Description("Compares substrings of two specified String objects, ignoring or honoring their case and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order.")]
        int Compare7(string strA, int indexA, string strB, int indexB, int length, bool ignoreCase, CultureInfo culture);

        [Description("Compares substrings of two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two substrings to each other in the sort order.")]
        int Compare8(string strA, int indexA, string strB, int indexB, int length, CultureInfo culture, CompareOptions options);

        // CompareOrdinal Overloads

        [Description("Compares two specified String objects by evaluating the numeric values of the corresponding Char objects in each string.")]
        int CompareOrdinal(string strA, string strB);

        [Description("Compares substrings of two specified String objects by evaluating the numeric values of the corresponding Char objects in each substring.")]
        int CompareOrdinal2(string strA, int indexA, string strB, int indexB, int length);

        [Description("Creates a new instance of String with the same value as a specified String.")]
        string Copy(string str);

        [Description("Determines whether two specified String objects have the same value.")]
        bool Equals(string a, string b);

        [Description("Determines whether two specified String objects have the same value. A parameter specifies the culture, case, and sort rules used in the comparison.")]
        bool Equals2(string a, string b, StringComparison comparisonType);


        // Format Overloads

        [Description("Replaces the format item in a specified string with the string representation of a corresponding object in a specified array.")]
        string Format(string pFormat, [In] ref object[] args);

        [Description("Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information.")]
        string Format2(IFormatProvider provider, string pFormat, [In] ref object[] args);

        [Description("Indicates whether the specified string is null or an empty string (\"\").")]
        bool IsNullOrEmpty(string value);

        [Description("Indicates whether a specified string is null, empty, or consists only of white-space characters.")]
        bool IsNullOrWhiteSpace(string value);
    }
}
