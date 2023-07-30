// https://www.thevbahelp.com/post/calling-c-sharp-code-from-vba-com-interop
// https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1

// Notes
// To fix member Parse3 passing arrays
// https://stackoverflow.com/questions/2027758/pass-an-array-from-vba-to-c-sharp-using-com-interop
// https://stackoverflow.com/a/2027776/10759363

using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("FC2EAD4C-E155-4A9B-989A-A6D93BAF4738")]
    [Description("Represents an instant in time, typically expressed as a date and time of day.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IComDateTime
    {
        // Constructors
        [Description("Initializes a new instance of the DateTime structure to a specified number of ticks and to Coordinated Universal Time (UTC) or local time.")]
        System.DateTime CreateFromTicks(long ticks, System.DateTimeKind kind = System.DateTimeKind.Unspecified);

        [Description("Initializes a new instance of the DateTime structure to the specified year, month, and day.")]
        System.DateTime CreateFromDate(int year, int month, int day);

        [Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and millisecond.")]
        System.DateTime CreateFromDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond = 0);

        [Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and Coordinated Universal Time (UTC) or local time.")]
        System.DateTime CreateFromDateTimeKind(int year, int month, int day, int hour, int minute, int second, DateTimeKind kind);

        [Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, millisecond, and Coordinated Universal Time (UTC) or local time.")]
        System.DateTime CreateFromDateTimeKind2(int year, int month, int day, int hour, int minute, int second, int millisecond, DateTimeKind kind);

        //Fields
        DateTime MaxValue 
        {
            [Description("Represents the largest possible value of DateTime. This field is read-only.")]
            get; 
        }

        DateTime MinValue 
        {
            [Description("Represents the smallest possible value of DateTime. This field is read-only.")]
            get; 
        }
        //DateTime UnixEpoch { get; }

        //Properties

        DateTime DateOnly 
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

        //int Nanosecond { get; }

        DateTime Now
        {
            [Description("Gets a DateTime object that is set to the current date and time on this computer, expressed as the local time.")]
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

        TimeSpan TimeOfDay 
        {
            [Description("Gets the time of day for this instance.")]
            get;
        }

        DateTime Today 
        {
            [Description("Gets the current date.")]
            get;
        }

        DateTime UtcNow 
        {
            [Description("Gets a DateTime object that is set to the current date and time on this computer, expressed as the Coordinated Universal Time (UTC).")]
            get;
        }

        int Year 
        {
            [Description("Gets the year component of the date represented by this instance.")]
            get;
        }

        // Methods

        [Description("Returns a new DateTime that adds the value of the specified TimeSpan to the value of this instance.")]
        DateTime Add(TimeSpan value);

        [Description("Returns a new DateTime that adds the specified number of days to the value of this instance.")]
        DateTime AddDays(double value);

        [Description("Returns a new DateTime that adds the specified number of hours to the value of this instance.")]
        DateTime AddHours(double value);
        //DateTimeTools AddMicoseconds(double value);

        [Description("Returns a new DateTime that adds the specified number of milliseconds to the value of this instance.")]
        DateTime AddMilliseconds(double value);

        [Description("Returns a new DateTime that adds the specified number of minutes to the value of this instance.")]
        DateTime AddMinutes(double value);

        [Description("Returns a new DateTime that adds the specified number of months to the value of this instance.")]
        DateTime AddMonths(int months);

        [Description("Returns a new DateTime that adds the specified number of seconds to the value of this instance.")]
        DateTime AddSeconds(double value);

        [Description("Returns a new DateTime that adds the specified number of ticks to the value of this instance.")]
        DateTime AddTicks(long value);

        [Description("Returns a new DateTime that adds the specified number of years to the value of this instance.")]
        DateTime AddYears(int value);

        [Description("Compares two instances of DateTime and returns an integer that indicates whether the first instance is earlier than, the same as, or later than the second instance.")]
        int Compare(DateTime t1, DateTime t2);

        [Description("Compares the value of this instance to a specified DateTime value and returns an integer that indicates whether this instance is earlier than, the same as, or later than the specified DateTime value.")]
        int CompareTo(DateTime value);

        [Description("Compares the value of this instance to a specified object that contains a specified DateTime value, and returns an integer that indicates whether this instance is earlier than, the same as, or later than the specified DateTime value.")]
        int CompareTo2(object value);

        [Description("Returns the number of days in the specified month and year.")]
        int DaysInMonth(int year, int month);

        [Description("Returns a value indicating whether the value of this instance is equal to the value of the specified DateTime instance.")]
        bool Equals(DateTime value);

        [Description("Returns a value indicating whether this instance is equal to a specified object.")]
        bool Equals2(object value);

        [Description("Returns a value indicating whether two DateTime instances have the same date and time value.")]
        bool Equals3(DateTime t1, DateTime t2);

        [Description("Deserializes a 64-bit binary value and recreates an original serialized DateTime object.")]
        DateTime FromBinary(long dateData);

        [Description("Converts the specified Windows file time to an equivalent local time.")]
        DateTime FromFileTime(long fileTime);

        [Description("Converts the specified Windows file time to an equivalent UTC time.")]
        DateTime FromFileTimeUtc(long fileTime);

        [Description("Returns a DateTime equivalent to the specified OLE Automation Date.")]
        DateTime FromOADate(double d);

        [Description("Converts the value of this instance to all the string representations supported by the standard date and time format specifiers.")]
        string[] GetDateTimeFormats();

        [Description("Returns the hash code for this instance.")]
        int GetHashCode();

        [Description("Indicates whether this instance of DateTime is within the daylight saving time range for the current time zone.")]
        bool IsDaylightSavingTime();

        [Description("Returns an indication whether the specified year is a leap year.")]
        bool IsLeapYear(int year);

        [Description("Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture.")]
        DateTime Parse(string s);

        [Description("Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture.")]
        DateTime Parse2(string s, IFormatProvider provider);

        [Description("Defines the formatting options that customize string parsing for some date and time parsing methods.  This enumeration supports a bitwise combination of its member values.")]
        DateTime Parse3(string s, IFormatProvider provider, GSystem.Globalization.DateTimeStyles styles);

        [Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified format and culture-specific format information. The format of the string representation must match the specified format exactly.")]
        DateTime ParseExact(string s, string format, IFormatProvider provider);

        [Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly or an exception is thrown.")]
        DateTime ParseExact2(string s, string format, IFormatProvider provider, GSystem.Globalization.DateTimeStyles style);

        [Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified array of formats, culture-specific format information, and style. The format of the string representation must match at least one of the specified formats exactly or an exception is thrown.")]
        DateTime ParseExact3(string s, [In] ref string[] formats, IFormatProvider provider, GSystem.Globalization.DateTimeStyles style);

        [Description("Creates a new DateTime object that has the same number of ticks as the specified DateTime, but is designated as either local time, Coordinated Universal Time (UTC), or neither, as indicated by the specified DateTimeKind value.")]
        DateTime SpecifyKind(DateTime value, DateTimeKind kind);

        [Description("Returns a new DateTime that subtracts the specified duration from the value of this instance.")]
        DateTime Subtract(TimeSpan value);

        [Description("Returns a new TimeSpan that subtracts the specified date and time from the value of this instance.")]
        TimeSpan Subtract2(DateTime value);

        [Description("Serializes the current DateTime object to a 64-bit binary value that subsequently can be used to recreate the DateTime object.")]
        long ToBinary();

        [Description("Converts the value of the current DateTime object to a Windows file time.")]
        long ToFileTime();

        [Description("Converts the value of the current DateTime object to a Windows file time.")] 
        long ToFileTimeUtc();

        [Description("Converts the value of the current DateTime object to local time.")]
        DateTime ToLocalTime();

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
        DateTime ToUniversalTime();

        [Description("Converts the specified string representation of a date and time to its DateTime equivalent and returns a value that indicates whether the conversion succeeded.")]
        bool TryParse(string s, out DateTime result);

        [Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified culture-specific format information and formatting style, and returns a value that indicates whether the conversion succeeded.")]
        bool TryParse2(string s, IFormatProvider provider, GSystem.Globalization.DateTimeStyles styles, out DateTime result);


        //Operators
        [Description("Adds a specified time interval to a specified date and time, yielding a new date and time.")]
        DateTime Addition(DateTime d, TimeSpan t);

        [Description("Determines whether two specified instances of DateTime are equal.")]
        bool Equality(DateTime d1, DateTime d2);

        [Description("Determines whether one specified DateTime is later than another specified DateTime.")]
        bool GreaterThan(DateTime t1, DateTime t2);

        [Description("Determines whether one specified DateTime represents a date and time that is the same as or later than another specified DateTime.")]
        bool GreaterThanOrEqual(DateTime t1, DateTime t2);

        [Description("Determines whether two specified instances of DateTime are not equal.")]
        bool Inequality(DateTime t1, DateTime t2);

        [Description("Determines whether one specified DateTime is earlier than another specified DateTime.")]
        bool LessThan(DateTime t1, DateTime t2);

        [Description("Determines whether one specified DateTime represents a date and time that is the same as or earlier than another specified DateTime.")]
        bool LessThanOrEqual(DateTime t1, DateTime t2);

        [Description("Subtracts a specified date and time from another specified date and time and returns a time interval.")]
        TimeSpan Subtraction(DateTime d1, DateTime d2);

        [Description("Subtracts a specified time interval from a specified date and time and returns a new date and time.")]
        DateTime Subtraction2(DateTime d, TimeSpan t);

    }
}
