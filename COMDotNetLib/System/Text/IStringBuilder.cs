// https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Guid("9EFD11A3-B1A8-409D-A562-3C62308CF5D2")]
    [Description("Represents a mutable string of characters. This class cannot be inherited.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IStringBuilder
    {
        int Capacity 
        {
            [Description("Gets or sets the maximum number of characters that can be contained in the memory allocated by the current instance.")]
            get;
            [Description("Gets or sets the maximum number of characters that can be contained in the memory allocated by the current instance.")]
            set; 
        }

        string this[int index] 
        {
            [Description("Gets or sets the character at the specified character position in this instance.")]
            get;
            [Description("Gets or sets the character at the specified character position in this instance.")]
            set;
        }

        int Length 
        {
            [Description("Gets or sets the length of the current StringBuilder object.")]
            get;
            [Description("Gets or sets the length of the current StringBuilder object.")]
            set;
        }

        int MaxCapacity 
        {
            [Description("Gets the maximum capacity of this instance.")]
            get;
        }

        // Methods

        [Description("Appends a copy of the specified string to this instance.")]
        StringBuilder Append(string value);

        [Description("Appends a copy of a specified substring to this instance.")] 
        StringBuilder Append(string value, int startIndex, int count);

        [Description("Appends the string representation of a specified Boolean value to this instance.")]
        StringBuilder Append(bool value);

        [Description("Appends the string representation of a specified 8-bit unsigned integer to this instance.")]
        StringBuilder Append(byte value);

        [Description("Appends the string representation of a specified 16-bit signed integer to this instance.")]
        StringBuilder Append(short value);

        [Description("Appends the string representation of a specified 32-bit signed integer to this instance.")]
        StringBuilder Append(int value);

        [Description("Appends the string representation of a specified 64-bit signed integer to this instance.")]
        StringBuilder Append(long value);

        [Description("Appends the string representation of a specified single-precision floating-point number to this instance.")]
        StringBuilder Append(float value);

        [Description("Appends the string representation of a specified double-precision floating-point number to this instance.")]   
        StringBuilder Append(double value);

        [Description("Appends the string representation of a specified decimal number to this instance.")] 
        StringBuilder Append(decimal value);

        [Description("Appends the string representation of a specified object to this instance.")]
        StringBuilder Append(object value);

        [Description("Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of a corresponding argument in a parameter array.")]
        StringBuilder AppendFormat(string format, [In] ref object[] args);

        [Description("Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of a single argument.")]
        StringBuilder AppendFormat(string format, object arg0);

        [Description("Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of either of two arguments.")]
        StringBuilder AppendFormat(string format, object arg0, object arg1);

        [Description("Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of either of three arguments.")]
        StringBuilder AppendFormat(string format, object arg0, object arg1, object arg2);

        [Description("Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of a corresponding argument in a parameter array using a specified format provider.")]
        StringBuilder AppendFormat(IFormatProvider provider, string pFormat, [In] ref object[] args);

        [Description("Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of a single argument using a specified format provider.")]
        StringBuilder AppendFormat(IFormatProvider provider, string format, object arg0);

        [Description("Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of either of two arguments using a specified format provider.")]
        StringBuilder AppendFormat(IFormatProvider provider, string format, object arg0, object arg1);

        [Description("Appends the string returned by processing a composite format string, which contains zero or more format items, to this instance. Each format item is replaced by the string representation of either of three arguments using a specified format provider.")]
        StringBuilder AppendFormat(IFormatProvider provider, string format, object arg0, object arg1, object arg2);

        [Description("Appends the default line terminator to the end of the current StringBuilder object.")]
        StringBuilder AppendLine();

        [Description("Appends a copy of the specified string followed by the default line terminator to the end of the current StringBuilder object.")]
        StringBuilder AppendLine(string value);

        [Description("Removes all characters from the current StringBuilder instance.")]
        StringBuilder Clear();

        [Description("Ensures that the capacity of this instance of StringBuilder is at least the specified value.")]
        int EnsureCapacity(int capacity);

        [Description("Returns a value indicating whether this instance is equal to a specified object.")]
        bool Equals(StringBuilder sb);


        [Description("Inserts one or more copies of a specified string into this instance at the specified character position.")]
        StringBuilder Insert(int index, string value);

        
        [Description("Inserts one or more copies of a specified string into this instance at the specified character position.")] 
        StringBuilder Insert(int index, string value, int count);

        [Description("Inserts the string representation of a Boolean value into this instance at the specified character position.")] 
        StringBuilder Insert(int index, bool value);

        
        [Description("Inserts the string representation of a specified 8-bit unsigned integer into this instance at the specified character position.")]
        StringBuilder Insert(int index, byte value);

        [Description("Inserts the string representation of a specified 16-bit signed integer into this instance at the specified character position.")]
        StringBuilder Insert(int index, short value);

        [Description("Inserts the string representation of a specified 32-bit signed integer into this instance at the specified character position.")] 
        StringBuilder Insert(int index, int value);

        [Description("Inserts the string representation of a 64-bit signed integer into this instance at the specified character position.")]
        StringBuilder Insert(int index, long value);

        [Description("Inserts the string representation of a single-precision floating point number into this instance at the specified character position.")]
        StringBuilder Insert(int index, float value);

        [Description("Inserts the string representation of a double-precision floating-point number into this instance at the specified character position.")]   
        StringBuilder Insert(int index, double value);

        [Description("Inserts the string representation of a decimal number into this instance at the specified character position.")] 
        StringBuilder Insert(int index, decimal value);

        [Description("Inserts the string representation of an object into this instance at the specified character position.")]
        StringBuilder Insert(int index, object value);

        [Description("Removes the specified range of characters from this instance.")]
        StringBuilder Remove(int startIndex, int length);

        [Description("Replaces all occurrences of a specified string in this instance with another specified string.")]
        StringBuilder Replace(string oldValue, string newValue);

        [Description("Replaces, within a substring of this instance, all occurrences of a specified string with another specified string.")]
        StringBuilder Replace(string oldValue, string newValue, int startIndex, int count);

        [Description("Converts the value of this instance to a String.")]
        string ToString();

        [Description("Converts the value of a substring of this instance to a String.")]
        string ToString(int startIndex, int length);

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

    }
}
