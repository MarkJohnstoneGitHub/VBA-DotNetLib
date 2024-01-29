// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.compareinfo?view=netframework-4.8.1

using GGlobalization = global::System.Globalization;
using System.Globalization;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("42189C56-2ED6-41BE-BB7A-E80A9CBCC783")]
    [Description("Implements a set of methods for culture-sensitive string comparisons.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICompareInfo
    {
        // Properties
        int LCID 
        {
            [Description("Gets the properly formed culture identifier for the current CompareInfo.")]
            get;
        }
        string Name 
        {
            [Description("Gets the name of the culture used for sorting operations by this CompareInfo object.")]
            get;
        }
        GGlobalization.SortVersion Version
        {
            [Description("Gets information about the version of Unicode used for comparing and sorting strings.")]
            get;
        }

        // Methods

        [Description("Compares two strings using the specified CompareOptions value.")]
        int Compare(string string1, string string2, CompareOptions options = CompareOptions.None);

        [Description("Compares the end section of a string with the end section of another string using the specified CompareOptions value.")]
        int Compare2(string string1, int offset1, string string2, int offset2, CompareOptions options = CompareOptions.None);

        [Description("Compares a section of one string with a section of another string using the specified CompareOptions value.")]
        int Compare3(string string1, int offset1, int length1, string string2, int offset2, int length2, CompareOptions options = CompareOptions.None);

        [Description("Determines whether the specified object is equal to the current CompareInfo object.")]
        bool Equals(object value);

        [Description("Serves as a hash function for the current CompareInfo for hashing algorithms and data structures, such as a hash table.")]
        int GetHashCode();

        [Description("Gets a SortKey object for the specified string using the specified CompareOptions value.")]
        GGlobalization.SortKey GetSortKey(string source, CompareOptions options = CompareOptions.None);

        //[Description("Searches for the specified substring and returns the zero-based index of the first occurrence within the entire source string using the specified CompareOptions value.")]
        //int IndexOf2(string source, string value, CompareOptions options = CompareOptions.None);

        [Description("Searches for the specified substring and returns the zero-based index of the first occurrence within the section of the source string that extends from the specified index to the end of the string using the specified CompareOptions value.")]
        int IndexOf(string source, string value, int startIndex = 0, CompareOptions options = CompareOptions.None);

        [Description("Searches for the specified substring and returns the zero-based index of the first occurrence within the section of the source string that starts at the specified index and contains the specified number of elements using the specified CompareOptions value.")]
        int IndexOf2(string source, string value, int startIndex, int count, CompareOptions options = CompareOptions.None);


        [Description("Determines whether the specified source string starts with the specified prefix using the specified CompareOptions value.")]
        bool IsPrefix(string source, string prefix, CompareOptions options = CompareOptions.None);

        [Description("Determines whether the specified source string ends with the specified suffix using the specified CompareOptions value.")]
        bool IsSuffix(string source, string suffix, CompareOptions options = CompareOptions.None);

        [Description("Searches for the specified substring and returns the zero-based index of the last occurrence within the section of the source string that extends from the beginning of the string to the specified index using the specified CompareOptions value.")]
        int LastIndexOf(string source, string value, int startIndex = 0, CompareOptions options = CompareOptions.None);

        [Description("Searches for the specified substring and returns the zero-based index of the last occurrence within the section of the source string that contains the specified number of elements and ends at the specified index using the specified CompareOptions value.")]
        int LastIndexOf2(string source, string value, int startIndex, int count, CompareOptions options = CompareOptions.None);

        [Description("Returns a string that represents the current CompareInfo object.")]
        string ToString();

    }
}
