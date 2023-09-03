// https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("4895B32B-0349-40E2-A4FB-E26C398E93E9")]
    [Description("Represents an instant in time, typically expressed as a date and time of day.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDateTimeOffset
    {
        // Properties

        DateTime Date
        {
            [Description("Gets a DateTime value that represents the date component of the current DateTimeOffset object.")]
            get;
        }

        DateTime DateTime
        {
            [Description("Gets a DateTime value that represents the date and time of the current DateTimeOffset object.")]
            get;
        }

        int Day
        {
            [Description("Gets the day of the month represented by the current DateTimeOffset object.")]
            get;
        }

        DayOfWeek DayOfWeek
        {
            [Description("Gets the day of the week represented by the current DateTimeOffset object.")]
            get;
        }

        int DayOfYear
        {
            [Description("Gets the day of the year represented by the current DateTimeOffset object.")]
            get;
        }

        int Hour
        {
            [Description("Gets the hour component of the time represented by the current DateTimeOffset object.")]
            get;
        }


        DateTime LocalDateTime
        {
            [Description("Gets a DateTime value that represents the local date and time of the current DateTimeOffset object.")]
            get;
        }

        int Millisecond
        {
            [Description("Gets the millisecond component of the time represented by the current DateTimeOffset object.")]
            get;
        }

        int Minute
        {
            [Description("Gets the minute component of the time represented by the current DateTimeOffset object.")]
            get;
        }

        int Month
        {
            [Description("Gets the month component of the date represented by the current DateTimeOffset object.")]
            get;
        }

        TimeSpan Offset
        {
            [Description("Gets the time's offset from Coordinated Universal Time (UTC).")]
            get;
        }

        int Second
        {
            [Description("Gets the second component of the clock time represented by the current DateTimeOffset object.")]
            get;
        }

        long Ticks
        {
            [Description("Gets the number of ticks that represents the date and time of the current DateTimeOffset object in clock time.")]
            get;
        }

        TimeSpan TimeOfDay
        {
            [Description("Gets the time of day for the current DateTimeOffset object.")]
            get;
        }

        DateTime UtcDateTime
        {
            [Description("Gets a DateTime value that represents the Coordinated Universal Time (UTC) date and time of the current DateTimeOffset object.")]
            get;
        }

        long UtcTicks
        {
            [Description("Gets the number of ticks that represents the date and time of the current DateTimeOffset object in Coordinated Universal Time (UTC).")]
            get;
        }

        int Year
        {
            [Description("Gets the year component of the date represented by the current DateTimeOffset object.")]
            get;
        }

        // Methods

        [Description("Returns a new DateTimeOffset object that adds a specified time interval to the value of this instance.")]
        DateTimeOffset Add(TimeSpan timeSpan);

        [Description("Returns a new DateTimeOffset object that adds a specified number of whole and fractional days to the value of this instance.")]
        DateTimeOffset AddDays(double pDays);

        [Description("Returns a new DateTimeOffset object that adds a specified number of whole and fractional hours to the value of this instance.")]
        DateTimeOffset AddHours(double pHours);

        [Description("Returns a new DateTimeOffset object that adds a specified number of milliseconds to the value of this instance.")]
        DateTimeOffset AddMilliseconds(double pMilliseconds);

        [Description("Returns a new DateTimeOffset object that adds a specified number of whole and fractional minutes to the value of this instance.")]
        DateTimeOffset AddMinutes(double pMinutes);

        [Description("Returns a new DateTimeOffset object that adds a specified number of months to the value of this instance.")]
        DateTimeOffset AddMonths(int pMonths);

        [Description("Returns a new DateTimeOffset object that adds a specified number of whole and fractional seconds to the value of this instance.")]
        DateTimeOffset AddSeconds(double pSeconds);

        [Description("Returns a new DateTimeOffset object that adds a specified number of ticks to the value of this instance.")]
        DateTimeOffset AddTicks(long pTicks);

        [Description("Returns a new DateTimeOffset object that adds a specified number of years to the value of this instance.")]
        DateTimeOffset AddYears(int pYears);

        [Description("Compares the current DateTimeOffset object to a specified DateTimeOffset object and indicates whether the current object is earlier than, the same as, or later than the second DateTimeOffset object.")]
        int CompareTo(DateTimeOffset other);

        [Description("Compares the value of the current DateTimeOffset object with another object of the same type.")]
        int CompareTo2(object value);

        [Description("Determines whether the current DateTimeOffset object represents the same point in time as a specified DateTimeOffset object.")]
        bool Equals(DateTimeOffset other);

        [Description("Determines whether a DateTimeOffset object represents the same point in time as a specified object.")]
        bool Equals2(object obj);

        [Description("Determines whether the current DateTimeOffset object represents the same time and has the same offset as a specified DateTimeOffset object.")]
        bool EqualsExact(DateTimeOffset other);

        [Description("Returns the hash code for the current DateTimeOffset object.")]
        int GetHashCode();

        [Description("Subtracts a DateTimeOffset value that represents a specific date and time from the current DateTimeOffset object.")]
        TimeSpan Subtract(DateTimeOffset value);

        [Description("Subtracts a specified time interval from the current DateTimeOffset object.")]
        DateTimeOffset Subtract2(TimeSpan value);

        [Description("Converts the value of the current DateTimeOffset object to a Windows file time.")]
        long ToFileTime();

        [Description("Converts the current DateTimeOffset object to a DateTimeOffset object that represents the local time.")]
        DateTimeOffset ToLocalTime();

        [Description("Converts the value of the current DateTimeOffset object to the date and time specified by an offset value.")]
        DateTimeOffset ToOffset(TimeSpan pOffset);

        [Description("Converts the value of the current DateTimeOffset object to its equivalent string representation.")]
        string ToString();

        [Description("Converts the value of the current DateTimeOffset object to its equivalent string representation using the specified format.")]
        string ToString2(string format);

        [Description("Converts the value of the current DateTimeOffset object to its equivalent string representation using the specified culture-specific formatting information.")]
        string ToString3(IFormatProvider formatProvider);

        [Description("Converts the value of the current DateTimeOffset object to its equivalent string representation using the specified format and culture-specific format information.")]
        string ToString4(string format, IFormatProvider formatProvider);

        [Description("Converts the current DateTimeOffset object to a DateTimeOffset value that represents the Coordinated Universal Time (UTC).")]
        DateTimeOffset ToUniversalTime();

        [Description("Returns the number of milliseconds that have elapsed since 1970-01-01T00:00:00.000Z.")]
        long ToUnixTimeMilliseconds();

        [Description("Returns the number of seconds that have elapsed since 1970-01-01T00:00:00Z.")]
        long ToUnixTimeSeconds();
    }
}
