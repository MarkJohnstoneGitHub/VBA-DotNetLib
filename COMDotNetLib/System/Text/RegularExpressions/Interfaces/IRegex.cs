// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("53D93876-6CEE-4369-9140-4035CEEE91A7")]
    [Description("Represents an immutable regular expression.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRegex
    {
        // Properties
        TimeSpan MatchTimeout 
        {
            [Description("Gets the time-out interval of the current instance.")]
            get;
        }

        RegexOptions Options 
        {
            [Description("Gets the options that were passed into the Regex constructor.")]
            get;
        }

        bool RightToLeft 
        {
            [Description("Gets a value that indicates whether the regular expression searches from right to left.")]
            get;
        }

        // Methods
        [Description("Returns an array of capturing group names for the regular expression.")]
        string[] GetGroupNames();

        [Description("Returns an array of capturing group numbers that correspond to group names in an array.")]
        int[] GetGroupNumbers();

        [Description("Gets the group name that corresponds to the specified group number.")]
        string GroupNameFromNumber(int i);

        [Description("Returns the group number that corresponds to the specified group name.")]
        int GroupNumberFromName(string name);

        [Description("Indicates whether the regular expression specified in the Regex constructor finds a match in a specified input string.")]
        bool IsMatch(string input);

        [Description("Indicates whether the regular expression specified in the Regex constructor finds a match in the specified input string, beginning at the specified starting position in the string.")]
        bool IsMatch2(string input, int startat);

        [Description("Searches the specified input string for the first occurrence of the regular expression specified in the Regex constructor.")]
        Match Match(string input);

        [Description("Searches the input string for the first occurrence of a regular expression, beginning at the specified starting position in the string.")]
        Match Match2(string input, int startat);

        [Description("Searches the input string for the first occurrence of a regular expression, beginning at the specified starting position and searching only the specified number of characters.")]
        Match Match3(string input, int beginning, int length);

        [Description("Searches the specified input string for all occurrences of a regular expression.")]
        MatchCollection Matches(string input);

        [Description("Searches the specified input string for all occurrences of a regular expression, beginning at the specified starting position in the string.")]
        MatchCollection Matches2(string input, int startat);

        [Description("In a specified input string, replaces all strings that match a regular expression pattern with a specified replacement string.")]
        string Replace(string input, string replacement);

        [Description("In a specified input substring, replaces a specified maximum number of strings that match a regular expression pattern with a specified replacement string.")]
        string Replace2(string input, string replacement, int count, int startat);

        //[Description("Splits an input string into an array of substrings at the positions defined by a regular expression pattern specified in the Regex constructor.")]
        //string[] Split2(string input);

        [Description("Splits an input string a specified maximum number of times into an array of substrings, at the positions defined by a regular expression specified in the Regex constructor.")]
        string[] Split(string input, int count = 0);

        [Description("Splits an input string a specified maximum number of times into an array of substrings, at the positions defined by a regular expression specified in the Regex constructor. The search for the regular expression pattern starts at a specified character position in the input string.")]
        string[] Split2(string input, int count, int startat);

        [Description("Returns the regular expression pattern that was passed into the Regex constructor.")]
        string ToString();

    }
}
