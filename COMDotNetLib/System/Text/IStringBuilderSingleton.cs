// https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Guid("D701B3EB-440B-4309-8EA2-41106969EF8C")]
    [Description("Represents a mutable string of characters. This class cannot be inherited.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStringBuilderSingleton
    {
        [Description("Initializes a new instance of the StringBuilder class using the specified string and capacity.")]
        StringBuilder Create(string value = "", int capacity = 16);

        [Description("Initializes a new instance of the StringBuilder class that starts with a specified capacity and can grow to a specified maximum.")]
        StringBuilder Create2(int capacity, int maxCapacity);

        [Description("Initializes a new instance of the StringBuilder class from the specified substring and capacity.")]
        StringBuilder Create3(string value, int startIndex, int length, int capacity);

    }
}
