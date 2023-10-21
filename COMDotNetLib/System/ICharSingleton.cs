// https://learn.microsoft.com/en-us/dotnet/api/system.char?view=netframework-4.8.1

using GGlobalization = global::System.Globalization;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("9EBC237C-645D-41DF-B1A5-2F85E20C09D9")]
    [Description("Represents a character as a UTF-16 code unit.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICharSingleton
    {
        [Description("Converts the specified Unicode code point into a UTF-16 encoded string.")]
        string ConvertFromUtf32(int utf32);

        [Description("Converts the value of a UTF-16 encoded character or surrogate pair at a specified position in a string into a Unicode code point.")]
        int ConvertToUtf32(string s, int index);

        [Description("Converts the numeric Unicode character at the specified position in a specified string to a double-precision floating point number.")]
        double GetNumericValue(string s, int index);

        [Description("Categorizes a Unicode character into a group identified by one of the UnicodeCategory values.")]
        GGlobalization.UnicodeCategory GetUnicodeCategory(string s, int index);

        [Description("Indicates whether the character at the specified position in a specified string is categorized as a control character.")]
        bool IsControl(string s, int index);

        [Description("Indicates whether the character at the specified position in a specified string is categorized as a decimal digit.")]
        bool IsDigit(string s, int index);

        [Description("Indicates whether the Char object at the specified position in a string is a high surrogate.")] 
        bool IsHighSurrogate(string s, int index);

        [Description("Indicates whether the character at the specified position in a specified string is categorized as a Unicode letter.")]
        bool IsLetter(string s, int index);

        [Description("Indicates whether the character at the specified position in a specified string is categorized as a letter or a decimal digit.")]
        bool IsLetterOrDigit(string s, int index);

        [Description("Indicates whether the character at the specified position in a specified string is categorized as a lowercase letter.")]
        bool IsLower(string s, int index);

        [Description("Indicates whether the Char object at the specified position in a string is a low surrogate.")]
        bool IsLowSurrogate(string s, int index);

        [Description("Indicates whether the character at the specified position in a specified string is categorized as a number.")]
        bool IsNumber(string s, int index);

        [Description("Indicates whether the character at the specified position in a specified string is categorized as a punctuation mark.")]
        bool IsPunctuation(string s, int index);

        [Description("Indicates whether the character at the specified position in a specified string is categorized as a separator character.")]
        bool IsSeparator(string s, int index);

        [Description("Indicates whether the character at the specified position in a specified string has a surrogate code unit.")]
        bool IsSurrogate(string s, int index);

        [Description("Indicates whether two adjacent Char objects at a specified position in a string form a surrogate pair.")]
        bool IsSurrogatePair(string s, int index);

        [Description("Indicates whether the character at the specified position in a specified string is categorized as a symbol character.")]
        bool IsSymbol(string s, int index);

        [Description("Indicates whether the character at the specified position in a specified string is categorized as an uppercase letter.")]
        bool IsUpper(string s, int index);

        [Description("Indicates whether the character at the specified position in a specified string is categorized as white space.")]
        bool IsWhiteSpace(string s, int index);



        //[Description("")]
    }
}
