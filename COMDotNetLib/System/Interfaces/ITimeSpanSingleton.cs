// https://learn.microsoft.com/en-us/dotnet/api/system.timespan?view=netframework-4.8

//Notes: https://www.geeksforgeeks.org/c-sharp-inheritance-in-interfaces/

using DotNetLib.System.Globalization;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using GSystem = global::System;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("C2F50F87-A7D4-4C32-A01B-EC750CE1B955")]
    [Description("Represents a time interval.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITimeSpanSingleton
    {
        //Constructors

        [Description("Initializes a new instance of the TimeSpan structure to the specified number of ticks.")]
        ITimeSpan CreateFromTicks(long ticks);

        [Description("Initializes a new instance of the TimeSpan structure to a specified number of hours, minutes, and seconds.")]
        ITimeSpan Create(int pHours, int pMinutes, int pSeconds);

        [Description("Initializes a new instance of the TimeSpan structure to a specified number of days, hours, minutes, and seconds.")]
        ITimeSpan Create2(int pDays, int pHours, int pMinutes, int pSeconds);

        [Description("Initializes a new instance of the TimeSpan structure to a specified number of days, hours, minutes, seconds, and milliseconds.")]
        ITimeSpan Create3(int pDays, int pHours, int pMinutes, int pSeconds, int pMilliseconds);

        // Fields

        ITimeSpan MaxValue 
        {
            [Description("Represents the maximum TimeSpan value. This field is read-only.")]
            get; 
        }

        ITimeSpan MinValue
        {
            [Description("Represents the minimum TimeSpan value. This field is read-only.")]
            get; 
        }

        long TicksPerDay 
        {
            [Description("Represents the number of ticks in 1 day. This field is constant.")]
            get; 
        }

        long TicksPerHour
        {
            [Description("Represents the number of ticks in 1 hour. This field is constant.")]
            get; 
        }

        long TicksPerMillisecond 
        {
            [Description("Represents the number of ticks in 1 millisecond. This field is constant.")]
            get;
        }

        long TicksPerMinute
        {
            [Description("Represents the number of ticks in 1 minute. This field is constant.")]
            get;
        }

        long TicksPerSecond 
        {
            [Description("Represents the number of ticks in 1 second.")]
            get;
        }

        ITimeSpan Zero
        {
            [Description("Represents the zero TimeSpan value. This field is read-only.")]
            get;
        }

        // Properties


        // Methods

        [Description("Compares two TimeSpan values and returns an integer that indicates whether the first value is shorter than, equal to, or longer than the second value.")]
        int Compare(TimeSpan t1, TimeSpan t2);

        [Description("Returns a value that indicates whether two specified instances of TimeSpan are equal.")]
        bool Equals3(TimeSpan t1, TimeSpan t2);

        [Description("Returns a TimeSpan that represents a specified number of days, where the specification is accurate to the nearest millisecond.")]
        ITimeSpan FromDays(double value);

        [Description("Returns a TimeSpan that represents a specified number of hours, where the specification is accurate to the nearest millisecond.")]
        ITimeSpan FromHours(double value);

        [Description("Returns a TimeSpan that represents a specified number of milliseconds.")]
        ITimeSpan FromMilliseconds(double value);

        [Description("Returns a TimeSpan that represents a specified number of minutes, where the specification is accurate to the nearest millisecond.")]
        ITimeSpan FromMinutes(double value);

        [Description("Returns a TimeSpan that represents a specified number of seconds, where the specification is accurate to the nearest millisecond.")]
        ITimeSpan FromSeconds(double value);

        [Description("Returns a TimeSpan that represents a specified time, where the specification is in units of ticks.")]
        ITimeSpan FromTicks(long value);

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent.")]
        ITimeSpan Parse(string s);

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified culture-specific format information.")]
        ITimeSpan Parse2(string input, GSystem.IFormatProvider formatProvider);

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly.")] 
        ITimeSpan ParseExact(string input, string format, IFormatProvider formatProvider);

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified array of format strings and culture-specific format information. The format of the string representation must match one of the specified formats exactly.")]
        ITimeSpan ParseExact2(string input, [In] ref string[] formats, IFormatProvider formatProvider);

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format, culture-specific format information, and styles. The format of the string representation must match the specified format exactly.")]
        ITimeSpan ParseExact3(string input, string format, GSystem.IFormatProvider formatProvider, TimeSpanStyles styles);

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified formats, culture-specific format information, and styles. The format of the string representation must match one of the specified formats exactly")]
        ITimeSpan ParseExact4(string input, [In] ref string[] formats, GSystem.IFormatProvider formatProvider, TimeSpanStyles styles);

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent and returns a value that indicates whether the conversion succeeded.")]
        bool TryParse(string s, out TimeSpan result);

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified culture-specific formatting information, and returns a value that indicates whether the conversion succeeded.")]
        bool TryParse2(string input, GSystem.IFormatProvider formatProvider, out TimeSpan result);

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly.")]
        bool TryParseExact(string input, string format, GSystem.IFormatProvider formatProvider, out TimeSpan result);

        [Description("Converts the specified string representation of a time interval to its TimeSpan equivalent by using the specified formats and culture-specific format information. The format of the string representation must match one of the specified formats exactly.")]
        bool TryParseExact2(string input, [In] ref string[] formats, GSystem.IFormatProvider formatProvider, out TimeSpan result);

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format, culture-specific format information and styles. The format of the string representation must match the specified format exactly.")]
        bool TryParseExact3(string input, string format, GSystem.IFormatProvider formatProvider, TimeSpanStyles styles, out TimeSpan result);

        [Description("Converts the specified string representation of a time interval to its TimeSpan equivalent by using the specified formats, culture-specific format information and styles. The format of the string representation must match one of the specified formats exactly.")]
        bool TryParseExact4(string input, [In] ref string[] formats, IFormatProvider formatProvider, TimeSpanStyles styles, out TimeSpan result);

        // Operators

        [Description("Adds two specified TimeSpan instances.")]
        TimeSpan Addition(TimeSpan t1, TimeSpan t2);

        [Description("Indicates whether two TimeSpan instances are equal.")]
        bool Equality(TimeSpan t1, TimeSpan t2);

        [Description("Indicates whether a specified TimeSpan is greater than another specified TimeSpan.")]
        bool GreaterThan(TimeSpan t1, TimeSpan t2);

        [Description("Indicates whether a specified TimeSpan is greater than or equal to another specified TimeSpan.")]
        bool GreaterThanOrEqual(TimeSpan t1, TimeSpan t2);

        [Description("Indicates whether two TimeSpan instances are not equal.")]
        bool Inequality(TimeSpan t1, TimeSpan t2);

        [Description("Indicates whether a specified TimeSpan is less than another specified TimeSpan.")]
        bool LessThan(TimeSpan t1, TimeSpan t2);

        [Description("Indicates whether a specified TimeSpan is less than or equal to another specified TimeSpan.")]
        bool LessThanOrEqual(TimeSpan t1, TimeSpan t2);

        [Description("Subtracts a specified TimeSpan from another specified TimeSpan.")]
        ITimeSpan Subtraction(TimeSpan t1, TimeSpan t2);

        [Description("Returns a TimeSpan whose value is the negated value of the specified instance.")]
        ITimeSpan UnaryNegation(TimeSpan ts);

        [Description("Returns the specified instance of TimeSpan.")]
        ITimeSpan UnaryPlus(TimeSpan ts);
    }
}
