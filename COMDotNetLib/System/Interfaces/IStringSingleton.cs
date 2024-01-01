// https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

using GSystem = global::System;
using GCollections = global::System.Collections;
using GGlobalization = global::System.Globalization;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using CultureInfo = DotNetLib.System.Globalization.CultureInfo;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("4F39D64D-9BD3-4AF7-A124-7A88364BE29F")]
    [Description("Represents text as a sequence of UTF-16 code units.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStringSingleton
    {
        // Factory Methods
        [Description("Initializes a new instance of the String class to the Unicode characters indicated in the specified string.")]
        String Create(string value);

        [Description("Initializes a new instance of the String class to the value indicated by a specified Unicode character repeated a specified number of times.")]
        String Create(string character, int count);

        [Description("Initializes a new instance of the String class to the value indicated by an string of Unicode characters, a starting character position within that array, and a length")]
        String Create(string value, int startIndex, int length);

        [Description("Initializes a new instance of the String class to the Unicode characters indicated in the specified string.")]
        String Create(String value);

        [Description("Initializes a new instance of the String class to the value indicated by a specified Unicode character repeated a specified number of times.")]
        String Create(String character, int count);

        [Description("Initializes a new instance of the String class to the value indicated by an string of Unicode characters, a starting character position within that array, and a length")]
        String Create(String value, int startIndex, int length);

        [Description("Initializes a new instance of the String class to the value indicated by an string of Unicode characters, converts any escaped characters in the input string.")]
        String CreateUnescape(string value);

        // Fields
        String EmptyString
        {
            [Description("Represents the empty string. This field is read-only.")]
            get;
        }

        // Compare Overloads

        [Description("Compares two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order.")]
        int Compare(String strA, String strB, bool ignoreCase = false);

        [Description("Compares two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order.")]
        int Compare2(String strA, String strB, GSystem.StringComparison comparisonType);

        [Description("Compares two specified String objects, ignoring or honoring their case, and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order.")]
        int Compare3(String strA, String strB, bool ignoreCase, CultureInfo culture);

        [Description("Compares two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two strings to each other in the sort order.")]
        int Compare4(String strA, String strB, CultureInfo culture, GGlobalization.CompareOptions options);

        [Description("Compares substrings of two specified String objects, ignoring or honoring their case, and returns an integer that indicates their relative position in the sort order.")]
        int Compare5(String strA, int indexA, String strB, int indexB, int length, bool ignoreCase = false);

        [Description("Compares substrings of two specified String objects using the specified rules, and returns an integer that indicates their relative position in the sort order.")]
        int Compare6(String strA, int indexA, String strB, int indexB, int length, GSystem.StringComparison comparisonType);

        [Description("Compares substrings of two specified String objects, ignoring or honoring their case and using culture-specific information to influence the comparison, and returns an integer that indicates their relative position in the sort order.")]
        int Compare7(String strA, int indexA, String strB, int indexB, int length, bool ignoreCase, CultureInfo culture);

        [Description("Compares substrings of two specified String objects using the specified comparison options and culture-specific information to influence the comparison, and returns an integer that indicates the relationship of the two substrings to each other in the sort order.")]
        int Compare8(String strA, int indexA, String strB, int indexB, int length, CultureInfo culture, GGlobalization.CompareOptions options);

        [Description("Compares two specified String objects by evaluating the numeric values of the corresponding Char objects in each string.")]
        int CompareOrdinal(String strA, String strB);

        [Description("Compares substrings of two specified String objects by evaluating the numeric values of the corresponding Char objects in each substring.")]
        int CompareOrdinal2(String strA, int indexA, String strB, int indexB, int length);


        [Description("Creates the string representation of a specified object.")]
        String Concat(object arg0);

        [Description("Concatenates the string representations of two specified objects.")]
        String Concat2(object arg0, object arg1);

        [Description("Concatenates the string representations of three specified objects.")]
        String Concat3(object arg0, object arg1, object arg2);

        [Description("Concatenates the string representations of the elements in a specified Object array.")]
        String Concat4([In] ref object[] values);

        [Description("Concatenates the members of a constructed IEnumerable<T> collection of type String.")]
        String Concat5(GCollections.IEnumerable stringValues);

        [Description("Concatenates the members of an IEnumerable<T> implementation.")]
        String Concat6(GCollections.IEnumerable values);

        [Description("Concatenates two specified instances of String.")]
        String Concat7(string str0, string str1);

        [Description("Concatenates three specified instances of String.")]
        String Concat8(string str0, string str1, string str2);

        [Description("Concatenates four specified instances of String.")]
        String Concat9(string str0, string str1, string str2, string str3);

        [Description("Concatenates the elements of a specified String array.")]
        String Concat10([In] ref string[] values);

        [Description("Concatenates two specified instances of String.")]
        String Concat12(String str0, String str1);

        [Description("Concatenates three specified instances of String.")]
        String Concat13(String str0, String str1, String str2);

        [Description("Concatenates four specified instances of String.")]
        String Concat14(String str0, String str1, String str2, String str3);

        [Description("Creates a new instance of String with the same value as a specified String.")]
        String Copy(string str);

        [Description("Creates a new instance of String with the same value as a specified String.")]
        String Copy(String str);

        [Description("Determines whether two specified String objects have the same value.")]
        bool Equals(String a, String b);

        [Description("Determines whether two specified String objects have the same value. A parameter specifies the culture, case, and sort rules used in the comparison.")]
        bool Equals2(String a, String b, GSystem.StringComparison comparisonType);

        [Description("Replaces the format item in a specified string with the string representation of a corresponding object in a specified array.")]
        String Format(string pFormat, [In] ref object[] args);

        [Description("Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information.")]
        String Format2(IFormatProvider provider, string pFormat, [In] ref object[] args);

        [Description("Replaces the format item in a specified string with the string representation of a corresponding object in a specified array.")]
        String Format3(String pFormat, [In] ref object[] args);

        [Description("Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information.")]
        String Format4(IFormatProvider provider, String pFormat, [In] ref object[] args);

        [Description("Indicates whether the specified string is null or an empty string (\"\").")]
        bool IsNullOrEmpty(String value);

        [Description("Indicates whether a specified string is null, empty, or consists only of white-space characters.")]
        bool IsNullOrWhiteSpace(String value);

        [Description("Concatenates all the elements of a string array, using the specified separator between each element.")]
        String Join(string separator, [In] ref string[] value);

        [Description("Concatenates the elements of an object array, using the specified separator between each element.")]
        String Join2(string separator, [In] ref object[] values);

        [Description("Concatenates the members of a constructed IEnumerable<T> collection of type String, using the specified separator between each member.")]
        String Join3(string separator, GCollections.IEnumerable stringValues);

        [Description("Concatenates the specified elements of a string array, using the specified separator between each element.")]
        String Join4(string separator, [In] ref string[] value, int startIndex, int count);

        // Extensions

        [Description("Initializes a new instance of the String class to the value indicated by an string of Unicode characters, converts any escaped characters in the input string.")]
        String Unescape(String value);

        [Description("Initializes a new instance of the String class to the value indicated by an string of Unicode characters, converts any escaped characters in the input string.")]
        String Unescape2(string value);

    }
}
