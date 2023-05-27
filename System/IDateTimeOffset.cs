using GSystem = global::System; // https://stackoverflow.com/questions/5681537/namespace-conflict-in-c-sharp
using System.Runtime.InteropServices;
using System.ComponentModel;
using System;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("818900B7-0353-45FA-8038-1C550219FD04")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]

    public interface IDateTimeOffset
    {
        // Constructors

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified DateTime value.")]
        DateTimeOffset CreateFromDateTime(DateTime dateTime);

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified DateTime value and offset.")]
        DateTimeOffset CreateFromDateTime2(DateTime dateTime, TimeSpan offset);

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, and offset.")]
        DateTimeOffset CreateFromDateTimeParts(int year, int month, int day, int hour, int minute, int second, TimeSpan offset);

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, millisecond, and offset.")]
        DateTimeOffset CreateFromDateTimeParts2(int year, int month, int day, int hour, int minute, int second, int millisecond, TimeSpan offset);

        DateTimeOffset CreateFromDateTimeParts3(int year, int month, int day, int hour, int minute, int second, int millisecond, GSystem.Globalization.Calendar calendar, TimeSpan offset);

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified number of ticks and offset.")]
        DateTimeOffset CreateFromTicks(long ticks, TimeSpan offset);


        // Fields

        [Description("Represents the greatest possible value of DateTimeOffset. This field is read-only.")]
        DateTimeOffset MaxValue { get; }

        [Description("Represents the earliest possible DateTimeOffset value. This field is read-only.")]
        DateTimeOffset MinValue { get; }

        // Properties

        [Description("Gets a DateTime value that represents the date component of the current DateTimeOffset object.")]
        DateTime Date { get; }

        [Description("Gets a DateTime value that represents the date and time of the current DateTimeOffset object.")]
        DateTime DateTime { get; }

        [Description("Gets the day of the month represented by the current DateTimeOffset object.")]
        int Day { get; }

        [Description("Gets the day of the week represented by the current DateTimeOffset object.")]
        DayOfWeek DayOfWeek { get; }

        [Description("Gets the day of the year represented by the current DateTimeOffset object.")]
        int DayOfYear { get; }

        [Description("Gets the hour component of the time represented by the current DateTimeOffset object.")]
        int Hour { get; }

        [Description("Gets a DateTime value that represents the local date and time of the current DateTimeOffset object.")]
        DateTime LocalDateTime { get; }

        [Description("Gets the millisecond component of the time represented by the current DateTimeOffset object.")]
        int Millisecond { get; }

        [Description("Gets the minute component of the time represented by the current DateTimeOffset object.")]
        int Minute { get; }

        [Description("Gets the month component of the date represented by the current DateTimeOffset object.")]
        int Month { get; }

        [Description("Gets a DateTimeOffset object that is set to the current date and time on the current computer, with the offset set to the local time's offset from Coordinated Universal Time (UTC).")]
        DateTimeOffset Now { get; }

        [Description("Gets the time's offset from Coordinated Universal Time (UTC).")]
        TimeSpan Offset { get; }

        [Description("Gets the second component of the clock time represented by the current DateTimeOffset object.")]
        int Second { get; }

        [Description("Gets the number of ticks that represents the date and time of the current DateTimeOffset object in clock time.")]
        long Ticks { get; }

        [Description("Gets the time of day for the current DateTimeOffset object.")]
        TimeSpan TimeOfDay { get; }

        [Description("Gets a DateTime value that represents the Coordinated Universal Time (UTC) date and time of the current DateTimeOffset object.")]
        DateTime UtcDateTime { get; }

        [Description("Gets a DateTimeOffset object whose date and time are set to the current Coordinated Universal Time (UTC) date and time and whose offset is Zero.")]
        DateTimeOffset UtcNow { get; }

        [Description("Gets the number of ticks that represents the date and time of the current DateTimeOffset object in Coordinated Universal Time (UTC).")]
        long UtcTicks { get; }

        [Description("Gets the year component of the date represented by the current DateTimeOffset object.")]
        int Year { get; }

        // Methods

        [Description("Returns a new DateTimeOffset object that adds a specified time interval to the value of this instance.")]
        DateTimeOffset Add(TimeSpan timeSpan);

        [Description("Returns a new DateTimeOffset object that adds a specified number of whole and fractional days to the value of this instance.")]
        DateTimeOffset AddDays(double days);

        [Description("Returns a new DateTimeOffset object that adds a specified number of milliseconds to the value of this instance.")]
        DateTimeOffset AddMilliseconds(double milliseconds);

        [Description("Returns a new DateTimeOffset object that adds a specified number of whole and fractional minutes to the value of this instance.")]
        DateTimeOffset AddMinutes(double minutes);

        [Description("Returns a new DateTimeOffset object that adds a specified number of months to the value of this instance.")]
        DateTimeOffset AddMonths(int months);

        [Description("Returns a new DateTimeOffset object that adds a specified number of whole and fractional seconds to the value of this instance.")] 
        DateTimeOffset AddSeconds(double seconds);

        [Description("Returns a new DateTimeOffset object that adds a specified number of ticks to the value of this instance.")]
        DateTimeOffset AddTicks(long ticks);

        [Description("Returns a new DateTimeOffset object that adds a specified number of years to the value of this instance.")]
        DateTimeOffset AddYears(int years);

        [Description("Compares two DateTimeOffset objects and indicates whether the first is earlier than the second, equal to the second, or later than the second.")]
        int Compare(DateTimeOffset first, DateTimeOffset second);

        [Description("Compares the current DateTimeOffset object to a specified DateTimeOffset object and indicates whether the current object is earlier than, the same as, or later than the second DateTimeOffset object.")]
        int CompareTo(DateTimeOffset other);

        [Description("Determines whether the current DateTimeOffset object represents the same point in time as a specified DateTimeOffset object.")]
        bool Equals(DateTimeOffset other);

        [Description("Determines whether a DateTimeOffset object represents the same point in time as a specified object.")]
        bool Equals2(object obj);

        [Description("Determines whether two specified DateTimeOffset objects represent the same point in time.")]
        bool Equals3(DateTimeOffset first, DateTimeOffset second);

        [Description("Determines whether the current DateTimeOffset object represents the same time and has the same offset as a specified DateTimeOffset object.")]
        bool EqualsExact(DateTimeOffset other);

        [Description("Converts the specified Windows file time to an equivalent local time.")]
        DateTimeOffset FromFileTime(long fileTime);

        [Description("Converts a Unix time expressed as the number of milliseconds that have elapsed since 1970-01-01T00:00:00Z to a DateTimeOffset value.")]
        DateTimeOffset FromUnixTimeMilliseconds(long milliseconds);

        [Description("Converts a Unix time expressed as the number of seconds that have elapsed since 1970-01-01T00:00:00Z to a DateTimeOffset value.")]
        DateTimeOffset FromUnixTimeSeconds(long seconds);

        [Description("Returns the hash code for the current DateTimeOffset object.")]
        int GetHashCode();

        [Description("Converts the specified string representation of a date, time, and offset to its DateTimeOffset equivalent.")]
        DateTimeOffset Parse(string input);

        [Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified culture-specific format information.")]
        DateTimeOffset Parse2(string input, IFormatProvider formatProvider);

        [Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified culture-specific format information and formatting style.")]
        DateTimeOffset Parse3(string input, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles);

        [Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly.")]
        DateTimeOffset ParseExact(string input, string format, IFormatProvider formatProvider);

        [Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly.")] 
        DateTimeOffset ParseExact2(string input, string format, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles);

        [Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified formats, culture-specific format information, and style. The format of the string representation must match one of the specified formats exactly.")]
        DateTimeOffset ParseExact3(string input, string[] formats, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles);

        [Description("Subtracts a DateTimeOffset value that represents a specific date and time from the current DateTimeOffset object.")]
        TimeSpan Subtract(DateTimeOffset value);

        [Description("Subtracts a specified time interval from the current DateTimeOffset object.")]
        DateTimeOffset Subtract2(TimeSpan value);

        // [Description("")]


    }
}
