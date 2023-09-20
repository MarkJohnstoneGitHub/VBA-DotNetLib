// https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1

using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.System.Globalization;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("FC2EAD4C-E155-4A9B-989A-A6D93BAF4738")]
    [Description("Represents an instant in time, typically expressed as a date and time of day.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDateTimeSingleton
    {
        // Constructors

        [Description("Initializes a new instance of the DateTime structure to a specified number of ticks and to Coordinated Universal Time (UTC) or local time.")]
        DateTime CreateFromTicks(long pTicks, System.DateTimeKind pKind = System.DateTimeKind.Unspecified);

        [Description("Initializes a new instance of the DateTime structure to the specified year, month, and day.")]
        DateTime CreateFromDate(int pYear, int pMonth, int pDay);

        [Description("Initializes a new instance of the DateTime structure to the specified year, month, and day for the specified calendar.")]
        DateTime CreateFromDate2(int pYear, int pMonth, int pDay, ICalendar calendar);

        [Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and millisecond.")]
        DateTime CreateFromDateTime(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, int pMillisecond = 0);

        [Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, and second for the specified calendar.")]
        DateTime CreateFromDateTime2(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, ICalendar calendar);

        [Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and millisecond for the specified calendar.")]
        DateTime CreateFromDateTime3(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, int pMillisecond, ICalendar calendar);

        [Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and Coordinated Universal Time (UTC) or local time.")]
        DateTime CreateFromDateTimeKind(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, DateTimeKind pKind);

        [Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, millisecond, and Coordinated Universal Time (UTC) or local time.")]
        DateTime CreateFromDateTimeKind2(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, int pMillisecond, DateTimeKind pKind);

        [Description("Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, millisecond, and Coordinated Universal Time (UTC) or local time for the specified calendar.")]
        DateTime CreateFromDateTimeKind3(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, int pMillisecond, ICalendar calendar, DateTimeKind pKind);

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
        DateTime Now
        {
            [Description("Gets a DateTime object that is set to the current date and time on this computer, expressed as the local time.")]
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

        [Description("Compares two instances of DateTime and returns an integer that indicates whether the first instance is earlier than, the same as, or later than the second instance.")]
        int Compare(DateTime t1, DateTime t2);

        [Description("Returns the number of days in the specified month and year.")]
        int DaysInMonth(int pYear, int pMonth);

        [Description("Returns a value indicating whether two DateTime instances have the same date and time value.")]
        bool Equals(DateTime t1, DateTime t2);

        [Description("Deserializes a 64-bit binary value and recreates an original serialized DateTime object.")]
        DateTime FromBinary(long dateData);

        [Description("Converts the specified Windows file time to an equivalent local time.")]
        DateTime FromFileTime(long fileTime);

        [Description("Converts the specified Windows file time to an equivalent UTC time.")]
        DateTime FromFileTimeUtc(long fileTime);

        [Description("Returns a DateTime equivalent to the specified OLE Automation Date.")]
        DateTime FromOADate(double d);

        [Description("Returns an indication whether the specified year is a leap year.")]
        bool IsLeapYear(int pYear);

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
        DateTime SpecifyKind(DateTime value, DateTimeKind pkind);

        [Description("Converts the specified string representation of a date and time to its DateTime equivalent and returns a value that indicates whether the conversion succeeded.")]
        bool TryParse(string s, out DateTime result);

        [Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified culture-specific format information and formatting style, and returns a value that indicates whether the conversion succeeded.")]
        bool TryParse2(string s, IFormatProvider provider, GSystem.Globalization.DateTimeStyles styles, out DateTime result);

        [Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified format, culture-specific format information, and style.")]
        bool TryParseExact(string s, string format, IFormatProvider provider, GSystem.Globalization.DateTimeStyles style, out DateTime result);

        [Description("Converts the specified string representation of a date and time to its DateTime equivalent using the specified array of formats, culture-specific format information, and style.")]
        bool TryParseExact2(string s, [In] ref string[] formats, IFormatProvider provider, GSystem.Globalization.DateTimeStyles style, out DateTime result);

        // Operators

        [Description("Adds a specified time interval to a specified date and time, yielding a new date and time.")]
        DateTime Addition(DateTime dt, TimeSpan ts);

        [Description("Determines whether two specified instances of DateTime are equal.")]
        bool Equality(DateTime dt1, DateTime dt2);

        [Description("Determines whether one specified DateTime is later than another specified DateTime.")]
        bool GreaterThan(DateTime dt1, DateTime dt2);

        [Description("Determines whether one specified DateTime represents a date and time that is the same as or later than another specified DateTime.")]
        bool GreaterThanOrEqual(DateTime dt1, DateTime dt2);

        [Description("Determines whether two specified instances of DateTime are not equal.")]
        bool Inequality(DateTime dt1, DateTime dt2);

        [Description("Determines whether one specified DateTime is earlier than another specified DateTime.")]
        bool LessThan(DateTime dt1, DateTime dt2);

        [Description("Determines whether one specified DateTime represents a date and time that is the same as or earlier than another specified DateTime.")]
        bool LessThanOrEqual(DateTime dt1, DateTime dt2);

        [Description("Subtracts a specified date and time from another specified date and time and returns a time interval.")]
        TimeSpan Subtraction(DateTime dt1, DateTime dt2);

        [Description("Subtracts a specified time interval from a specified date and time and returns a new date and time.")]
        DateTime Subtraction2(DateTime dt, TimeSpan ts);
    }
}
