// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.group?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("5073EC3F-1D79-4EE4-82D0-495BB681A277")]
    [Description("Represents the results from a single capturing group.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IGroup : ICapture
    {
        // Properties

        [Description("Gets a collection of all the captures matched by the capturing group, in innermost-leftmost-first order (or innermost-rightmost-first order if the regular expression is modified with the RightToLeft option). The collection may have zero or more items.")]
        CaptureCollection Captures 
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

        [Description("Returns the name of the capturing group represented by the current instance.")]
        string Name 
        { 
            get;
        }

        [Description("Gets a value indicating whether the match is successful.")]
        bool Success 
        { 
            get;
        }

        new string Value 
        {
            [Description("Gets the captured substring from the input string.\r\n\r\n(Inherited from Capture)")]
            get;
        }

        //Todo /ReadOnlySpan<char> ValueSpan { get; }
        //ReadOnlySpan<char> ValueSpan { get; }


        // Methods

        [Description("Determines whether the specified object is equal to the current object.\r\n\r\n(Inherited from Object)")]
        new bool Equals(object obj);

        [Description("Serves as the default hash function.\r\n\r\n(Inherited from Object)")]
        new int GetHashCode();

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

        [Description("Retrieves the captured substring from the input string by calling the Value property.\r\n\r\n(Inherited from Capture)")]
        new string ToString();
    }
}
