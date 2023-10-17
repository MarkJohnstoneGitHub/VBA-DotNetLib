// https://learn.microsoft.com/en-us/dotnet/api/system.stringcomparer?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("2622664A-9458-42B6-9C90-296851126660")]
    [Description("Represents a string comparison operation that uses specific case and culture-based or ordinal comparison rules.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStringComparer
    {

        [Description("When overridden in a derived class, compares two objects and returns an indication of their relative sort order.")]
        int Compare(object x, object y);

        [Description("When overridden in a derived class, compares two strings and returns an indication of their relative sort order.")]
        int Compare(string x, string y);

        [Description("Determines whether the specified object is equal to the current object.\r\n\r\n(Inherited from Object)")]
        bool Equals(object obj);

        [Description("When overridden in a derived class, indicates whether two objects are equal.")]
        bool Equals(object x, object y);

        [Description("When overridden in a derived class, indicates whether two strings are equal.")]
        bool Equals(string x, string y);

        [Description("Serves as the default hash function.\r\n\r\n(Inherited from Object)")]
        int GetHashCode();

        [Description("When overridden in a derived class, gets the hash code for the specified object.")]
        int GetHashCode(object obj);

        [Description("When overridden in a derived class, gets the hash code for the specified string.")]
        int GetHashCode(string obj);

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

        [Description("Returns a string that represents the current object.\r\n\r\n(Inherited from Object)")]
        string ToString();

    }
}
