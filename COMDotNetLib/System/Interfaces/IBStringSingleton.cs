// https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

using GSystem = global::System;
using GCollections = global::System.Collections;
using CultureInfo = DotNetLib.System.Globalization.CultureInfo;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("2A3FC380-AC92-4455-A564-A65C8DCD98BF")]
    [Description("Represents text as a sequence of UTF-16 code units.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IBStringSingleton
    {
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

        [Description("Compares two specified String objects by evaluating the numeric values of the corresponding Char objects in each string.")]
        int CompareOrdinal(string strA, string strB);

        [Description("Compares substrings of two specified String objects by evaluating the numeric values of the corresponding Char objects in each substring.")]
        int CompareOrdinal(string strA, int indexA, string strB, int indexB, int length);

        [Description("Concatenates the elements of a specified String array.")]
        string Concat([In] ref string[] values);

        [Description("Concatenates two specified instances of String.")]
        string Concat2(string str0, string str1);

        [Description("Concatenates three specified instances of String.")]
        string Concat3(string str0, string str1, string str2);

        [Description("Concatenates four specified instances of String.")]
        string Concat4(string str0, string str1, string str2, string str3);

        [Description("Concatenates the members of a constructed IEnumerable<T> collection of type String.")]
        string Concat5(GCollections.IEnumerable stringValues);

        [Description("Concatenates the string representations of the elements in a specified Object array.")]
        string Concat6([In] ref object[] values);

        [Description("Creates the string representation of a specified object.")]
        string Concat7(object arg0);

        [Description("Concatenates the string representations of two specified objects.")]
        string Concat8(object arg0, object arg1);

        [Description("Concatenates the string representations of three specified objects.")]
        string Concat9(object arg0, object arg1, object arg2);

        [Description("Concatenates the members of an IEnumerable<T> implementation.")]
        string Concat10(GCollections.IEnumerable values);

        [Description("Returns a value indicating whether a specified substring occurs within this string, using the specified comparison rules.")]
        bool Contains(string str, string substring, GSystem.StringComparison comparisonType = GSystem.StringComparison.Ordinal);

        [Description("Creates a new instance of String with the same value as a specified String.")]
        string Copy(string str);

        [Description("Determines whether the end of this string instance matches the specified string when compared using the specified comparison option.")]
        bool EndsWith(string str, string substring, GSystem.StringComparison comparisonType = StringComparison.CurrentCulture);

        [Description("Determines whether the end of this string instance matches the specified string when compared using the specified culture.")]
        bool EndsWith2(string str, string substring, bool ignoreCase, CultureInfo culture);

        [Description("Determines whether two specified String objects have the same value.")]
        bool Equals(string a, string b);

        [Description("Determines whether two specified String objects have the same value. A parameter specifies the culture, case, and sort rules used in the comparison.")]
        bool Equals2(string a, string b, StringComparison comparisonType);

        [Description("Replaces the format item in a specified string with the string representation of a corresponding object in a specified array.")]
        string Format(string pFormat, [In] ref object[] args);

        [Description("Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information.")]
        string Format2(IFormatProvider provider, string pFormat, [In] ref object[] args);

        [Description("Indicates whether the specified string is null or an empty string (\"\").")]
        bool IsNullOrEmpty(string value);

        [Description("Indicates whether a specified string is null, empty, or consists only of white-space characters.")]
        bool IsNullOrWhiteSpace(string value);

        [Description("Concatenates all the elements of a string array, using the specified separator between each element.")]
        string Join(string separator, [In] ref string[] value);

        [Description("Concatenates the elements of an object array, using the specified separator between each element.")]
        string Join2(string separator, [In] ref object[] values);

        [Description("Concatenates the specified elements of a string array, using the specified separator between each element.")]
        string Join4(string separator, [In] ref string[] value, int startIndex, int count);

        [Description("Concatenates the members of a constructed IEnumerable<T> collection of type String, using the specified separator between each member.")]
        string Join3(string separator, GCollections.IEnumerable stringValues);

        [Description("Splits a string into substrings based on specified delimiting characters and options.")]
        string[] Split(string str, string separators, StringSplitOptions options = StringSplitOptions.None);

        [Description("Splits a string into a maximum number of substrings based on specified delimiting characters and, optionally, options.")]
        string[] Split2(string str, string separator, int count, StringSplitOptions options = StringSplitOptions.None);

        [Description("Splits a string into substrings based on a specified delimiting string and, optionally, options")]
        string[] Split3(string str, [In] ref string[] separator, StringSplitOptions options);

        [Description("Splits a string into a maximum number of substrings based on specified delimiting strings and, optionally, options.")]
        string[] Split4(string str, [In] ref string[] separator, int count, StringSplitOptions options);

        [Description("Determines whether the beginning of this string instance matches the specified string when compared using the specified comparison option.")]
        bool StartsWith(string str, string substring, GSystem.StringComparison comparisonType = StringComparison.CurrentCulture);

        [Description("Determines whether the beginning of this string instance matches the specified string when compared using the specified culture.")]
        bool StartsWith2(string str, string substring, bool ignoreCase, CultureInfo culture);

        //Extensions

        [Description("Initializes a new instance of the String class to the value indicated by an string of Unicode characters, converts any escaped characters in the input string.")]
        string Unescape(string value);

    }
}
