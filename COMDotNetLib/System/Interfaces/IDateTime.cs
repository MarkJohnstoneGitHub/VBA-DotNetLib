//https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("95FDDD19-48C8-4FAC-8655-CD06F864F98E")]
    [Description("Represents an instant in time, typically expressed as a date and time of day.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDateTime
    {
        //Properties

        //Note: Renamed from Date due to VBA reserved word
        IDateTime DateOnly
        {
            [Description("Gets the date component of this instance.")]
            get;
        }

        int Day
        {
            [Description("Gets the day of the month represented by this instance.")]
            get;
        }

        DayOfWeek DayOfWeek
        {
            [Description("Gets the day of the week represented by this instance.")]
            get;
        }

        int DayOfYear
        {
            [Description("Gets the day of the year represented by this instance.")]
            get;
        }

        int Hour
        {
            [Description("Gets the hour component of the date represented by this instance.")]
            get;
        }

        DateTimeKind Kind
        {
            [Description("Gets a value that indicates whether the time represented by this instance is based on local time, Coordinated Universal Time (UTC), or neither.")]
            get;
        }

        //int Microsecond { get; }

        int Millisecond
        {
            [Description("Gets the milliseconds component of the date represented by this instance.")]
            get;
        }

        int Minute
        {
            [Description("Gets the minute component of the date represented by this instance.")]
            get;
        }

        int Month
        {
            [Description("Gets the month component of the date represented by this instance.")]
            get;
        }

        int Second
        {
            [Description("Gets the seconds component of the date represented by this instance.")]
            get;
        }

        long Ticks
        {
            [Description("Gets the number of ticks that represent the date and time of this instance.")]
            get;
        }

        ITimeSpan TimeOfDay
        {
            [Description("Gets the time of day for this instance.")]
            get;
        }

        int Year
        {
            [Description("Gets the year component of the date represented by this instance.")]
            get;
        }

        // Methods

        [Description("Returns a new DateTime that adds the value of the specified TimeSpan to the value of this instance.")]
        IDateTime Add(TimeSpan value);

        [Description("Returns a new DateTime that adds the specified number of days to the value of this instance.")]
        IDateTime AddDays(double value);

        [Description("Returns a new DateTime that adds the specified number of hours to the value of this instance.")]
        IDateTime AddHours(double value);

        //DateTimeTools AddMicoseconds(double value);

        [Description("Returns a new DateTime that adds the specified number of milliseconds to the value of this instance.")]
        IDateTime AddMilliseconds(double value);

        [Description("Returns a new DateTime that adds the specified number of minutes to the value of this instance.")]
        IDateTime AddMinutes(double value);

        [Description("Returns a new DateTime that adds the specified number of months to the value of this instance.")]
        IDateTime AddMonths(int months);

        [Description("Returns a new DateTime that adds the specified number of seconds to the value of this instance.")]
        IDateTime AddSeconds(double value);

        [Description("Returns a new DateTime that adds the specified number of ticks to the value of this instance.")]
        IDateTime AddTicks(long value);

        [Description("Returns a new DateTime that adds the specified number of years to the value of this instance.")]
        IDateTime AddYears(int value);

        [Description("Compares the value of this instance to a specified DateTime value and returns an integer that indicates whether this instance is earlier than, the same as, or later than the specified DateTime value.")]
        int CompareTo(DateTime value);

        [Description("Compares the value of this instance to a specified object that contains a specified DateTime value, and returns an integer that indicates whether this instance is earlier than, the same as, or later than the specified DateTime value.")]
        int CompareTo2(object value);

        [Description("Returns a value indicating whether the value of this instance is equal to the value of the specified DateTime instance.")]
        bool Equals(DateTime value);

        [Description("Returns a value indicating whether this instance is equal to a specified object.")]
        bool Equals2(object value);

        [Description("Converts the value of this instance to all the string representations supported by the standard date and time format specifiers.")]
        string[] GetDateTimeFormats();

        [Description("Returns the hash code for this instance.")]
        int GetHashCode();

        [Description("Indicates whether this instance of DateTime is within the daylight saving time range for the current time zone.")]
        bool IsDaylightSavingTime();

        [Description("Returns a new DateTime that subtracts the specified duration from the value of this instance.")]
        IDateTime Subtract(TimeSpan value);

        [Description("Returns a new TimeSpan that subtracts the specified date and time from the value of this instance.")]
        ITimeSpan Subtract2(DateTime value);

        [Description("Serializes the current DateTime object to a 64-bit binary value that subsequently can be used to recreate the DateTime object.")]
        long ToBinary();

        [Description("Converts the value of the current DateTime object to a Windows file time.")]
        long ToFileTime();

        [Description("Converts the value of the current DateTime object to a Windows file time.")]
        long ToFileTimeUtc();

        [Description("Converts the value of the current DateTime object to local time.")]
        IDateTime ToLocalTime();

        [Description("Converts the value of the current DateTime object to its equivalent long date string representation.")]
        string ToLongDateString();

        [Description("Converts the value of the current DateTime object to its equivalent long time string representation.")]
        string ToLongTimeString();

        [Description("Converts the value of this instance to the equivalent OLE Automation date.")]
        double ToOADate();

        [Description("Converts the value of the current DateTime object to its equivalent short date string representation.")]
        string ToShortDateString();

        [Description("Converts the value of the current DateTime object to its equivalent short time string representation.")]
        string ToShortTimeString();

        [Description("Converts the value of the current DateTime object to its equivalent string representation using the formatting conventions of the current culture.")]
        string ToString();

        [Description("Converts the value of the current DateTime object to its equivalent string representation using the specified format and the formatting conventions of the current culture.")]
        string ToString2(string format);

        [Description("Converts the value of the current DateTime object to its equivalent string representation using the specified culture-specific format information.")]
        string ToString3(IFormatProvider provider);

        [Description("Converts the value of the current DateTime object to its equivalent string representation using the specified format and culture-specific format information.")]
        string ToString4(string format, IFormatProvider provider);

        [Description("Converts the value of the current DateTime object to Coordinated Universal Time (UTC).")]
        IDateTime ToUniversalTime();
    }
}
