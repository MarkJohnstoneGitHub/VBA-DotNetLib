using System;
// https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

using GSystem = global::System;
using GText = global::System.Text;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.System.Globalization;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("9F7F9539-A372-4DFB-975D-003A4C6C3EC9")]
    [Description("Represents text as a sequence of UTF-16 code units.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IString
    {
        string WrappedString
        {
            [Description("Wrapped string object")]
            get;
        }

        int Length 
        {
            [Description("Gets the number of characters in the current String object.")]
            get;
        }

        // Methods
        [Description("Returns a reference to this instance of String.")]
        object Clone();

        [Description("Compares this instance with a specified String object and indicates whether this instance precedes, follows, or appears in the same position in the sort order as the specified string.")]
        int CompareTo(String strB);

        [Description("Compares this instance with a specified String object and indicates whether this instance precedes, follows, or appears in the same position in the sort order as the specified string.")]
        int CompareTo2(string strB);

        [Description("Compares this instance with a specified Object and indicates whether this instance precedes, follows, or appears in the same position in the sort order as the specified Object.")]
        int CompareTo3(object obj);

        [Description("Returns a value indicating whether a specified substring occurs within this string.")]
        bool Contains(String value);

        [Description("Returns a value indicating whether a specified substring occurs within this string.")]
        bool Contains2(string value);

        [Description("Determines whether the end of this string instance matches the specified string.")]
        bool EndsWith(String value);

        [Description("Determines whether the end of this string instance matches the specified string when compared using the specified comparison option.")]
        bool EndsWith2(String value, GSystem.StringComparison comparisonType);

        [Description("Determines whether the end of this string instance matches the specified string when compared using the specified culture.")]
        bool EndsWith3(String value, bool ignoreCase, CultureInfo culture);

        [Description("Determines whether the end of this string instance matches the specified string.")]
        bool EndsWith4(string value);

        [Description("Determines whether the end of this string instance matches the specified string when compared using the specified comparison option.")]
        bool EndsWith5(string value, GSystem.StringComparison comparisonType);

        [Description("Determines whether the end of this string instance matches the specified string when compared using the specified culture.")]
        bool EndsWith6(string value, bool ignoreCase, CultureInfo culture);

        [Description("Determines whether this instance and another specified String object have the same value.")]
        bool Equals(String value);

        [Description("Determines whether this string and a specified String object have the same value. A parameter specifies the culture, case, and sort rules used in the comparison.")]
        bool Equals2(String value, GSystem.StringComparison comparisonType);

        [Description("Determines whether this instance and another specified String object have the same value.")]
        bool Equals3(string value);

        [Description("Determines whether this string and a specified String object have the same value. A parameter specifies the culture, case, and sort rules used in the comparison.")]
        bool Equals4(string value, GSystem.StringComparison comparisonType);

        [Description("Returns the hash code for this string.")]
        int GetHashCode();

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

        [Description("Returns the TypeCode for the String class.")]
        GSystem.TypeCode GetTypeCode();

        [Description("Reports the zero-based index of the first occurrence of the specified string in this instance.")]
        int IndexOf(String value);

        [Description("Reports the zero-based index of the first occurrence of the specified string in the current String object. A parameter specifies the type of search to use for the specified string.")]
        int IndexOf2(String value, GSystem.StringComparison comparisonType);

        [Description("Reports the zero-based index of the first occurrence of the specified string in this instance. The search starts at a specified character position.")]
        int IndexOf3(String value, int startIndex);

        [Description("Reports the zero-based index of the first occurrence of the specified string in the current String object. Parameters specify the starting search position in the current string and the type of search to use for the specified string.")]
        int IndexOf4(String value, int startIndex, StringComparison comparisonType);

        [Description("Reports the zero-based index of the first occurrence of the specified string in this instance. The search starts at a specified character position and examines a specified number of character positions.")]
        int IndexOf5(String value, int startIndex, int count);

        [Description("Reports the zero-based index of the first occurrence of the specified string in the current String object. Parameters specify the starting search position in the current string, the number of characters in the current string to search, and the type of search to use for the specified string.")]
        int IndexOf6(String value, int startIndex, int count, GSystem.StringComparison comparisonType);

        [Description("Reports the zero-based index of the first occurrence of the specified string in this instance.")]
        int IndexOf7(string value);

        [Description("Reports the zero-based index of the first occurrence of the specified string in the current String object. A parameter specifies the type of search to use for the specified string.")]
        int IndexOf8(string value, GSystem.StringComparison comparisonType);

        [Description("Reports the zero-based index of the first occurrence of the specified string in this instance. The search starts at a specified character position.")]
        int IndexOf9(string value, int startIndex);

        [Description("Reports the zero-based index of the first occurrence of the specified string in the current String object. Parameters specify the starting search position in the current string and the type of search to use for the specified string.")]
        int IndexOf10(string value, int startIndex, StringComparison comparisonType);

        [Description("Reports the zero-based index of the first occurrence of the specified string in this instance. The search starts at a specified character position and examines a specified number of character positions.")]
        int IndexOf11(string value, int startIndex, int count);

        [Description("Reports the zero-based index of the first occurrence of the specified string in the current String object. Parameters specify the starting search position in the current string, the number of characters in the current string to search, and the type of search to use for the specified string.")]
        int IndexOf12(string value, int startIndex, int count, GSystem.StringComparison comparisonType);

        [Description("Reports the zero-based index of the first occurrence in this instance of any character in a specified array of Unicode characters.")]
        int IndexOfAny(String anyOf);

        [Description("Reports the zero-based index of the first occurrence in this instance of any character in a specified array of Unicode characters. The search starts at a specified character position.")]
        int IndexOfAny2(String anyOf, int startIndex);

        [Description("Reports the zero-based index of the first occurrence in this instance of any character in a specified array of Unicode characters. The search starts at a specified character position and examines a specified number of character positions.")]
        int IndexOfAny3(String anyOf, int startIndex, int count);

        [Description("Reports the zero-based index of the first occurrence in this instance of any character in a specified array of Unicode characters.")]
        int IndexOfAny4(string anyOf);

        [Description("Reports the zero-based index of the first occurrence in this instance of any character in a specified array of Unicode characters. The search starts at a specified character position.")]
        int IndexOfAny5(string anyOf, int startIndex);

        [Description("Reports the zero-based index of the first occurrence in this instance of any character in a specified array of Unicode characters. The search starts at a specified character position and examines a specified number of character positions.")]
        int IndexOfAny6(string anyOf, int startIndex, int count);

        [Description("Returns a new string in which a specified string is inserted at a specified index position in this instance.")]
        String Insert(int startIndex, String value);

        [Description("Returns a new string in which a specified string is inserted at a specified index position in this instance.")]
        String Insert2(int startIndex, string value);

        [Description("Returns a new VBA string in which a specified string is inserted at a specified index position in this instance.")]
        string InsertBStr(int startIndex, string value);

        [Description("Indicates whether this string is in the specified Unicode normalization form.")]
        bool IsNormalized(GText.NormalizationForm normalizationForm = GText.NormalizationForm.FormC);

        [Description("Reports the zero-based index position of the last occurrence of a specified string within this instance.")]
        int LastIndexOf(String value);

        [Description("Reports the zero-based index of the last occurrence of a specified string within the current String object. A parameter specifies the type of search to use for the specified string.")]
        int LastIndexOf4(String value, GSystem.StringComparison comparisonType);

        [Description("Reports the zero-based index position of the last occurrence of a specified string within this instance. The search starts at a specified character position and proceeds backward toward the beginning of the string.")]
        int LastIndexOf2(String value, int startIndex);

        [Description("Reports the zero-based index position of the last occurrence of a specified string within this instance. The search starts at a specified character position and proceeds backward toward the beginning of the string for a specified number of character positions.")]
        int LastIndexOf3(String value, int startIndex, int count);

        [Description("Reports the zero-based index of the last occurrence of a specified string within the current String object. The search starts at a specified character position and proceeds backward toward the beginning of the string. A parameter specifies the type of comparison to perform when searching for the specified string.")]
        int LastIndexOf5(String value, int startIndex, StringComparison comparisonType);

        [Description("Reports the zero-based index position of the last occurrence of a specified string within this instance. The search starts at a specified character position and proceeds backward toward the beginning of the string for the specified number of character positions. A parameter specifies the type of comparison to perform when searching for the specified string.")]
        int LastIndexOf6(String value, int startIndex, int count, GSystem.StringComparison comparisonType);

        [Description("Reports the zero-based index position of the last occurrence of a specified string within this instance.")]
        int LastIndexOf7(string value);

        [Description("Reports the zero-based index of the last occurrence of a specified string within the current String object. A parameter specifies the type of search to use for the specified string.")]
        int LastIndexOf11(string value, int startIndex, StringComparison comparisonType);

        [Description("Reports the zero-based index of the last occurrence of a specified string within the current String object. A parameter specifies the type of search to use for the specified string.")]
        int LastIndexOf10(string value, GSystem.StringComparison comparisonType);

        [Description("Reports the zero-based index position of the last occurrence of a specified string within this instance. The search starts at a specified character position and proceeds backward toward the beginning of the string.")]
        int LastIndexOf8(string value, int startIndex);

        [Description("Reports the zero-based index position of the last occurrence of a specified string within this instance. The search starts at a specified character position and proceeds backward toward the beginning of the string for a specified number of character positions.")]
        int LastIndexOf9(string value, int startIndex, int count);

        [Description("Reports the zero-based index position of the last occurrence of a specified string within this instance. The search starts at a specified character position and proceeds backward toward the beginning of the string for the specified number of character positions. A parameter specifies the type of comparison to perform when searching for the specified string.")]
        int LastIndexOf12(string value, int startIndex, int count, GSystem.StringComparison comparisonType);

        [Description("Reports the zero-based index position of the last occurrence in this instance of one or more characters specified in a Unicode string.")]
        int LastIndexOfAny(String anyOf);

        [Description("Reports the zero-based index position of the last occurrence in this instance of one or more characters specified in a Unicode string. The search starts at a specified character position and proceeds backward toward the beginning of the string for a specified number of character positions.")]
        int LastIndexOfAny2(String anyOf, int startIndex);

        [Description("Reports the zero-based index position of the last occurrence in this instance of one or more characters specified in a Unicode string. The search starts at a specified character position and proceeds backward toward the beginning of the string.")]
        int LastIndexOfAny3(String anyOf, int startIndex, int count);

        [Description("Reports the zero-based index position of the last occurrence in this instance of one or more characters specified in a Unicode string.")]
        int LastIndexOfAny4(string anyOf);

        [Description("Reports the zero-based index position of the last occurrence in this instance of one or more characters specified in a Unicode string. The search starts at a specified character position and proceeds backward toward the beginning of the string.")]
        int LastIndexOfAny5(string anyOf, int startIndex);

        [Description("Reports the zero-based index position of the last occurrence in this instance of one or more characters specified in a Unicode string. The search starts at a specified character position and proceeds backward toward the beginning of the string for a specified number of character positions.")]
        int LastIndexOfAny6(string anyOf, int startIndex, int count);


        //[Description("")]
        //String Normalize();

        [Description("Returns a new string whose textual value is the same as this string, but whose binary representation is in Unicode normalization form C.")]
        String Normalize(GText.NormalizationForm normalizationForm = GText.NormalizationForm.FormC);

        //[Description("")]
        //string NormalizeBStr();

        [Description("Returns a new string whose textual value is the same as this string, but whose binary representation is in Unicode normalization form C.")]
        string NormalizeBStr(GText.NormalizationForm normalizationForm = GText.NormalizationForm.FormC);

        //[Description("Returns a new string that right-aligns the characters in this instance by padding them with spaces on the left, for a specified total length.")]
        //String PadLeft(int totalWidth);

        [Description("Returns a new string that right-aligns the characters in this instance by padding them on the left with spaces or a specified Unicode character, for a specified total length.")]
        String PadLeft(int totalWidth, string paddingChar = " ");

        //[Description("Returns a new string that right-aligns the characters in this instance by padding them with spaces on the left, for a specified total length.")]
        //string PadLeftBStr(int totalWidth);

        [Description("Returns a new string that right-aligns the characters in this instance by padding them on the left with spaces or a specified Unicode character, for a specified total length.")]
        string PadLeftBStr(int totalWidth, string paddingChar = " ");

        //[Description("Returns a new string that left-aligns the characters in this string by padding them with spaces on the right, for a specified total length.")]
        //String PadRight(int totalWidth);

        [Description("Returns a new string that left-aligns the characters in this string by padding them on the right with spaces or specified Unicode character, for a specified total length.")]
        String PadRight(int totalWidth, string paddingChar = " ");

        //[Description("Returns a new string that left-aligns the characters in this string by padding them with spaces on the right, for a specified total length.")]
        //string PadRightBStr(int totalWidth);

        [Description("Returns a new string that left-aligns the characters in this string by padding them on the right with spaces or a specified Unicode character, for a specified total length.")]
        string PadRightBStr(int totalWidth, string paddingChar = " ");

        [Description("Returns a new string in which all the characters in the current instance, beginning at a specified position and continuing through the last position, have been deleted.")]
        String Remove(int startIndex);

        [Description("Returns a new string in which a specified number of characters in the current instance beginning at a specified position have been deleted.")]
        String Remove2(int startIndex, int count);

        [Description("Returns a new string in which all the characters in the current instance, beginning at a specified position and continuing through the last position, have been deleted.")]
        string RemoveBStr(int startIndex);

        [Description("Returns a new string in which a specified number of characters in the current instance beginning at a specified position have been deleted.")]
        string RemoveBStr2(int startIndex, int count);


        [Description("Returns a new string in which all occurrences of a specified string in the current instance are replaced with another specified string")]
        String Replace(String oldValue, String newValue);

        [Description("Returns a new string in which all occurrences of a specified string in the current instance are replaced with another specified string")]
        string ReplaceBStr(string oldValue, string newValue);


        //[Description("Splits a string into substrings based on specified delimiting characters.")]
        //string[] Split2(string separator);

        [Description("Splits a string into substrings based on specified delimiting characters and options.")]
        string[] Split(string separator, StringSplitOptions options = StringSplitOptions.None);

        [Description("Splits a string into a maximum number of substrings based on specified delimiting characters and, optionally, options.")]
        string[] Split2(string separator, int count, StringSplitOptions options = StringSplitOptions.None);

        [Description("Splits a string into substrings based on a specified delimiting string and, optionally, options")]
        string[] Split3(string[] separator, StringSplitOptions options);

        [Description("Splits a string into a maximum number of substrings based on specified delimiting strings and, optionally, options.")]
        string[] Split4(string[] separator, int count, StringSplitOptions options);


        [Description("Splits a string into substrings based on specified delimiting characters.")]
        Array SplitStringArray(string separator);

        [Description("Splits a string into substrings based on specified delimiting characters and options.")]
        Array SplitStringArray2(string separator, StringSplitOptions options);

        [Description("Splits a string into a maximum number of substrings based on specified delimiting characters and, optionally, options.")]
        Array SplitStringArray3(string separator, int count, StringSplitOptions options = StringSplitOptions.None);

        [Description("Splits a string into substrings based on a specified delimiting string and, optionally, options")]
        Array SplitStringArray4(string[] separator, StringSplitOptions options);

        [Description("Splits a string into a maximum number of substrings based on specified delimiting strings and, optionally, options.")]
        Array SplitStringArray5(string[] separator, int count, StringSplitOptions options);


        [Description("Determines whether the beginning of this string instance matches the specified string.")]
        bool StartsWith(String value);

        [Description("Determines whether the beginning of this string instance matches the specified string when compared using the specified comparison option.")]
        bool StartsWith2(String value, GSystem.StringComparison comparisonType);

        [Description("Determines whether the beginning of this string instance matches the specified string when compared using the specified culture.")]
        bool StartsWith3(String value, bool ignoreCase, CultureInfo culture);

        [Description("Determines whether the beginning of this string instance matches the specified string.")]
        bool StartsWith4(string value);

        [Description("Determines whether the beginning of this string instance matches the specified string when compared using the specified comparison option.")]
        bool StartsWith5(string value, GSystem.StringComparison comparisonType);
        [Description("Determines whether the beginning of this string instance matches the specified string when compared using the specified culture.")]
        bool StartsWith6(string value, bool ignoreCase, CultureInfo culture);

        [Description("Retrieves a substring from this instance. The substring starts at a specified character position and continues to the end of the string.")]
        String Substring(int startIndex);

        [Description("Retrieves a substring from this instance. The substring starts at a specified character position and has a specified length.")]
        String Substring2(int startIndex, int length);

        [Description("Retrieves a substring from this instance. The substring starts at a specified character position and continues to the end of the string.")]
        string SubstringBStr(int startIndex);

        [Description("Retrieves a substring from this instance. The substring starts at a specified character position and has a specified length.")]
        string SubstringBStr2(int startIndex, int length);

        [Description("Returns a copy of this string converted to lowercase.")]
        String ToLower();

        [Description("Returns a copy of this string converted to lowercase, using the casing rules of the specified culture.")]
        String ToLower2(CultureInfo culture);

        [Description("Returns a copy of this string converted to lowercase.")]
        string ToLowerBStr();

        [Description("Returns a copy of this string converted to lowercase, using the casing rules of the specified culture.")]
        string ToLowerBStr2(CultureInfo culture);


        [Description("Returns a copy of this String object converted to lowercase using the casing rules of the invariant culture.")]
        String ToLowerInvariant();

        [Description("Returns a copy of this String object converted to lowercase using the casing rules of the invariant culture.")]
        string ToLowerInvariantBStr();

        [Description("Returns this instance of String; no actual conversion is performed.")]
        string ToString();


        [Description("Returns a copy of this string converted to uppercase.")]
        String ToUpper();

        [Description("Returns a copy of this string converted to uppercase, using the casing rules of the specified culture.")]
        String ToUpper2(CultureInfo culture);

        [Description("Returns a copy of this string converted to uppercase.")]
        string ToUpperBStr();

        [Description("Returns a copy of this string converted to uppercase, using the casing rules of the specified culture.")]
        string ToUpperBStr2(CultureInfo culture);

        
        [Description("Returns a copy of this String object converted to uppercase using the casing rules of the invariant culture.")]
        String ToUpperInvariant();

        [Description("Returns a copy of this String object converted to uppercase using the casing rules of the invariant culture.")]
        string ToUpperInvariantBStr();

        [Description("Removes all leading and trailing white-space characters from the current string.")]
        String Trim();

        [Description("Removes all leading and trailing occurrences of a set of characters specified in an string from the current string.")]
        String Trim2(String trimChars);

        [Description("Removes all leading and trailing occurrences of a set of characters specified in an string from the current string.")]
        String Trim3(string trimChars);

        [Description("Removes all leading and trailing white-space characters from the current string.")]
        string TrimBStr();

        [Description("Removes all leading and trailing occurrences of a set of characters specified in an array from the current string.")]
        string TrimBStr2(string trimChars);

        [Description("Removes all the trailing occurrences of a set of characters specified in an array from the current string.")]
        String TrimEnd(String trimChars);

        [Description("Removes all the trailing occurrences of a set of characters specified in an array from the current string.")]
        String TrimEnd2(string trimChars);

        [Description("Removes all the trailing occurrences of a set of characters specified in an array from the current string.")]
        string TrimEndBStr(string trimChars);

        [Description("Removes all the leading occurrences of a set of characters specified in an array from the current string.")]
        String TrimStart(String trimChars);

        [Description("Removes all the leading occurrences of a set of characters specified in an array from the current string.")]
        String TrimStart2(string trimChars);

        [Description("Removes all the leading occurrences of a set of characters specified in an array from the current string.")]
        string TrimStartBStr(string trimChars);

    }
}
