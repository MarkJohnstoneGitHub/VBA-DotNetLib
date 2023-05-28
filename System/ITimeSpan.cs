using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.System.Globalization;

namespace DotNetLib.System
{
    // https://learn.microsoft.com/en-us/dotnet/api/system.timespan?view=netframework-4.8

    [ComVisible(true)]
    [Guid("C2F50F87-A7D4-4C32-A01B-EC750CE1B955")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITimeSpan
    {
        //Constructors

        [Description("Initializes a new instance of the TimeSpan structure to the specified number of ticks.")]
        TimeSpan CreateFromTicks(long ticks);

        [Description("Initializes a new instance of the TimeSpan structure to a specified number of hours, minutes, and seconds.")]
        TimeSpan Create(int hours, int minutes, int seconds);

        [Description("Initializes a new instance of the TimeSpan structure to a specified number of days, hours, minutes, and seconds.")]
        TimeSpan Create2(int days, int hours, int minutes, int seconds);

        [Description("Initializes a new instance of the TimeSpan structure to a specified number of days, hours, minutes, seconds, and milliseconds.")]
        TimeSpan Create3(int days, int hours, int minutes, int seconds, int milliseconds);

        // Fields

        [Description("Represents the maximum TimeSpan value. This field is read-only.")]
        TimeSpan MaxValue { get; }

        [Description("Represents the minimum TimeSpan value. This field is read-only.")]
        TimeSpan MinValue { get; }

        [Description("Represents the number of ticks in 1 day. This field is constant.")]
        long TicksPerDay { get; }

        [Description("Represents the number of ticks in 1 hour. This field is constant.")]
        long TicksPerHour { get; }

        [Description("Represents the number of ticks in 1 millisecond. This field is constant.")]
        long TicksPerMillisecond { get; }

        [Description("Represents the number of ticks in 1 minute. This field is constant.")]
        long TicksPerMinute { get; }

        [Description("Represents the number of ticks in 1 second.")]
        long TicksPerSecond { get; }

        [Description("Represents the zero TimeSpan value. This field is read-only.")]
        TimeSpan Zero { get; }

        // Properties

        [Description("Gets the days component of the time interval represented by the current TimeSpan structure.")]
        int Days { get; }

        [Description("Gets the hours component of the time interval represented by the current TimeSpan structure.")]
        int Hours { get; }

        [Description("Gets the milliseconds component of the time interval represented by the current TimeSpan structure.")]
        int Milliseconds { get; }

        [Description("Gets the minutes component of the time interval represented by the current TimeSpan structure.")]
        int Minutes { get; }

        [Description("Gets the seconds component of the time interval represented by the current TimeSpan structure.")]
        int Seconds { get; }

        [Description("Gets the number of ticks that represent the value of the current TimeSpan structure.")]
        long Ticks { get; }

        [Description("Gets the value of the current TimeSpan structure expressed in whole and fractional days.")]
        double TotalDays { get; }

        [Description("Gets the value of the current TimeSpan structure expressed in whole and fractional hours.")]
        double TotalHours { get; }

        [Description("Gets the value of the current TimeSpan structure expressed in whole and fractional milliseconds.")]
        double TotalMilliseconds { get; }

        [Description("Gets the value of the current TimeSpan structure expressed in whole and fractional minutes.")]
        double TotalMinutes { get; }

        [Description("Gets the value of the current TimeSpan structure expressed in whole and fractional seconds.")]
        double TotalSeconds { get; }

        [Description("Returns a new TimeSpan object whose value is the sum of the specified TimeSpan object and this instance.")]
        TimeSpan Add(TimeSpan ts);

        [Description("Compares two TimeSpan values and returns an integer that indicates whether the first value is shorter than, equal to, or longer than the second value.")]
        int Compare(TimeSpan t1, TimeSpan t2);

        [Description("Compares this instance to a specified TimeSpan object and returns an integer that indicates whether this instance is shorter than, equal to, or longer than the TimeSpan object.")]
        int CompareTo(TimeSpan value);

        [Description("Compares this instance to a specified object and returns an integer that indicates whether this instance is shorter than, equal to, or longer than the specified object.")]
        int CompareTo2(object value);

        [Description("Returns a new TimeSpan object whose value is the absolute value of the current TimeSpan object.")]
        TimeSpan Duration();

        [Description("Returns a value indicating whether this instance is equal to a specified TimeSpan object.")]
        bool Equals(TimeSpan obj);

        [Description("Returns a value indicating whether this instance is equal to a specified object.")]
        bool Equals2(object value);

        [Description("Returns a value that indicates whether two specified instances of TimeSpan are equal.")]
        bool Equals3(TimeSpan t1, TimeSpan t2);

        [Description("Returns a TimeSpan that represents a specified number of days, where the specification is accurate to the nearest millisecond.")]
        TimeSpan FromDays(double value);

        [Description("Returns a TimeSpan that represents a specified number of hours, where the specification is accurate to the nearest millisecond.")]
        TimeSpan FromHours(double value);

        [Description("Returns a TimeSpan that represents a specified number of milliseconds.")]
        TimeSpan FromMilliseconds(double value);

        [Description("Returns a TimeSpan that represents a specified number of minutes, where the specification is accurate to the nearest millisecond.")]
        TimeSpan FromMinutes(double value);

        [Description("Returns a TimeSpan that represents a specified number of seconds, where the specification is accurate to the nearest millisecond.")]
        TimeSpan FromSeconds(double value);

        [Description("Returns a TimeSpan that represents a specified time, where the specification is in units of ticks.")]
        TimeSpan FromTicks(long value);

        [Description("Returns a hash code for this instance.")]
        int GetHashCode();

        [Description("Returns a new TimeSpan object whose value is the negated value of this instance.")]
        TimeSpan Negate();

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent.")]
        TimeSpan Parse(string s);

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified culture-specific format information.")]
        TimeSpan Parse2(string input, GSystem.IFormatProvider formatProvider);

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly.")] 
        TimeSpan ParseExact(string input, string format, IFormatProvider formatProvider);

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified array of format strings and culture-specific format information. The format of the string representation must match one of the specified formats exactly.")]
        TimeSpan ParseExact2(string input, string[] formats, IFormatProvider formatProvider);

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified format, culture-specific format information, and styles. The format of the string representation must match the specified format exactly.")]
        TimeSpan ParseExact3(string input, string format, GSystem.IFormatProvider formatProvider, TimeSpanStyles styles);

        [Description("Converts the string representation of a time interval to its TimeSpan equivalent by using the specified formats, culture-specific format information, and styles. The format of the string representation must match one of the specified formats exactly")]
        TimeSpan ParseExact4(string input, string[] formats, GSystem.IFormatProvider formatProvider, TimeSpanStyles styles);

        [Description("Returns a new TimeSpan object whose value is the difference between the specified TimeSpan object and this instance.")]
        TimeSpan Subtract(TimeSpan ts);

        [Description("Converts the value of the current TimeSpan object to its equivalent string representation.")]
        string ToString();

        [Description("Converts the value of the current TimeSpan object to its equivalent string representation by using the specified format.")]
        string ToString2(string format);

        [Description("Converts the value of the current TimeSpan object to its equivalent string representation by using the specified format and culture-specific formatting information.")]
        string ToString3(string format, GSystem.IFormatProvider formatProvider);


        [Description("")]
        bool TryParse(string s, out TimeSpan result);

        [Description("")]
        bool TryParse2(string input, GSystem.IFormatProvider formatProvider, out TimeSpan result);


    }
}
