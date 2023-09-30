// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.match?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("A7154879-E9DA-43F3-944B-9F4A237BD15C")]
    [Description("Represents the results from a single regular expression match.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IMatch : IGroup, ICapture
    {
        new CaptureCollection Captures
        {
            [Description("Gets a collection of all the captures matched by the capturing group, in innermost-leftmost-first order (or innermost-rightmost-first order if the regular expression is modified with the RightToLeft option). The collection may have zero or more items.\r\n\r\n(Inherited from Group)")]
            get;
        }

        [Description("Gets a collection of groups matched by the regular expression.")]
        GroupCollection Groups
        {
            get;
        }

        new int Index
        {
            [Description("The position in the original string where the first character of the captured substring is found.\r\n\r\n(Inherited from Capture)")]
            get;
        }

        new int Length
        {
            [Description("Gets the length of the captured substring.\r\n\r\n(Inherited from Capture)")]
            get;
        }

        [Description("Returns the name of the capturing group represented by the current instance.\r\n\r\n(Inherited from Group)")]
        new string Name
        {
            get;
        }

        [Description("Gets a value indicating whether the match is successful.\r\n\r\n(Inherited from Group)")]
        new bool Success
        {
            get;
        }

        new string Value
        {
            [Description("Gets the captured substring from the input string.\r\n\r\n(Inherited from Capture)")]
            get;
        }

        //ReadOnlySpan<char> ValueSpan { get; }


        // Methods

        [Description("Serves as the default hash function.\r\n\r\n(Inherited from Object)")]
        new int GetHashCode();

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        new Type GetType();

        [Description("Returns a new Match object with the results for the next match, starting at the position at which the last match ended (at the character after the last matched character).")]
        Match NextMatch();

        [Description("Returns the expansion of the specified replacement pattern.")]
        string Result(string replacement);

        [Description("Retrieves the captured substring from the input string by calling the Value property.\r\n\r\n(Inherited from Capture)")]
        new string ToString();

    }
}
