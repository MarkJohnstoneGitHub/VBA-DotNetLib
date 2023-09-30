// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.capture?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("1A4FC5EB-D853-4E65-BDD8-876B1FF4AE9F")]
    [Description("Represents the results from a single successful subexpression capture.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICapture
    {
        // Properties
        int Index 
        {
            [Description("The position in the original string where the first character of the captured substring is found.")]
            get; 
        }

        int Length
        {
            [Description("Gets the length of the captured substring.")]
            get;
        }

        string Value 
        {
            [Description("Gets the captured substring from the input string.")]
            get;
        }

        // Methods

        [Description("Determines whether the specified object is equal to the current object.\r\n\r\n(Inherited from Object)")]
        bool Equals(object obj);

        [Description("Serves as the default hash function.\r\n\r\n(Inherited from Object)")]
        int GetHashCode();

        [Description("Retrieves the captured substring from the input string by calling the Value property.")]
        string ToString();
    }
}
