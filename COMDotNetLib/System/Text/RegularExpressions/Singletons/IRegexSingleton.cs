// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("8D42ED1E-623C-4704-B008-4AEDBE5A36FD")]
    [Description("Represents an immutable regular expression.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRegexSingleton
    {
        // Factory Methods

        [Description("Initializes a new instance of the Regex class for the specified regular expression, with options that modify the pattern.")]
        Regex Create(string pattern, RegexOptions options = RegexOptions.None);

        [Description("Initializes a new instance of the Regex class for the specified regular expression, with options that modify the pattern and a value that specifies how long a pattern matching method should attempt a match before it times out.")]
        Regex Create2(string pattern, RegexOptions options, TimeSpan matchTimeout);

        // Properties
        int CacheSize 
        {
            [Description("Gets or sets the maximum number of entries in the current static cache of compiled regular expressions.")]
            get;
            [Description("Gets or sets the maximum number of entries in the current static cache of compiled regular expressions.")]
            set;
        }

        TimeSpan InfiniteMatchTimeout
        {
            [Description("Specifies that a pattern-matching operation should not time out.")]
            get;
        }

        // Methods

        [Description("Escapes a minimal set of characters by replacing them with their escape codes. This instructs the regular expression engine to interpret these characters literally rather than as metacharacters.")]
        string Escape(string str);

        [Description("Indicates whether the specified regular expression finds a match in the specified input string, using the specified matching options.")]
        bool IsMatch(string input, string pattern, RegexOptions options = RegexOptions.None);

        [Description("Indicates whether the specified regular expression finds a match in the specified input string, using the specified matching options and time-out interval.")]
        bool IsMatch2(string input, string pattern, RegexOptions options, TimeSpan matchTimeout);

        [Description("Searches the input string for the first occurrence of the specified regular expression, using the specified matching options.")]
        Match Match(string input, string pattern, RegexOptions options = RegexOptions.None);

        [Description("Searches the input string for the first occurrence of the specified regular expression, using the specified matching options and time-out interval")]
        Match Match2(string input, string pattern, RegexOptions options, TimeSpan matchTimeout);

        [Description("Searches the specified input string for all occurrences of a specified regular expression, using the specified matching options.")]
        MatchCollection Matches(string input, string pattern, RegexOptions options = RegexOptions.None);

        [Description("Searches the specified input string for all occurrences of a specified regular expression, using the specified matching options and time-out interval.")]
        MatchCollection Matches2(string input, string pattern, RegexOptions options, TimeSpan matchTimeout);

        [Description("In a specified input string, replaces all strings that match a specified regular expression with a specified replacement string. Specified options modify the matching operation.")]
        string Replace(string input, string pattern, string replacement, RegexOptions options = RegexOptions.None);

        [Description("In a specified input string, replaces all strings that match a specified regular expression with a specified replacement string. Additional parameters specify options that modify the matching operation and a time-out interval if no match is found.")]
        string Replace2(string input, string pattern, string replacement, RegexOptions options, TimeSpan matchTimeout);

        [Description("Splits an input string into an array of substrings at the positions defined by a specified regular expression pattern. Specified options modify the matching operation.")]
        string[] Split(string input, string pattern, RegexOptions options = RegexOptions.None);

        [Description("Splits an input string into an array of substrings at the positions defined by a specified regular expression pattern. Additional parameters specify options that modify the matching operation and a time-out interval if no match is found.")]
        string[] Split2(string input, string pattern, RegexOptions options, TimeSpan matchTimeout);

        [Description("Converts any escaped characters in the input string.")]
        string Unescape(string str);

    }
}
