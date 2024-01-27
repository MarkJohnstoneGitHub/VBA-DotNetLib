// https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Description("Represents a mutable string of characters. This class cannot be inherited.")]
    [Guid("11597AC9-1EFE-4964-8CA4-904BD8F05875")]
    [ProgId("DotNetLib.System.Text.StringBuilderSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IStringBuilderSingleton))]
    public class StringBuilderSingleton : IStringBuilderSingleton
    {
        // Factory Methods
        public StringBuilder Create(string value = "", int capacity = 16)
        {
            return new StringBuilder(value, capacity);
        }

        public StringBuilder Create2(int capacity, int maxCapacity)
        {
            return new StringBuilder(capacity, maxCapacity);
        }

        public StringBuilder Create3(string value, int startIndex, int length, int capacity)
        {
            return new StringBuilder(value, startIndex, length, capacity);
        }


    }

}
