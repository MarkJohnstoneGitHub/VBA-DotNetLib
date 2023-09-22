// https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset?view=netframework-4.8.1

using GSystem = global::System; // https://stackoverflow.com/questions/5681537/namespace-conflict-in-c-sharp
using System.Runtime.InteropServices;
using System.ComponentModel;
using System;
using DotNetLib.System.Globalization;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("818900B7-0353-45FA-8038-1C550219FD04")]
    [Description("Represents a point in time, typically expressed as a date and time of day, relative to Coordinated Universal Time (UTC).")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]

    public interface IDateTimeOffsetSingleton
    {
        // Constructors

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified DateTime value.")]
        DateTimeOffset CreateFromDateTime(DateTime pDateTime);

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified DateTime value and offset.")]
        DateTimeOffset CreateFromDateTime2(DateTime pDateTime, TimeSpan pOffset);

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, and offset.")]
        DateTimeOffset CreateFromDateTimeParts(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, TimeSpan pOffset);

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, millisecond, and offset.")]
        DateTimeOffset CreateFromDateTimeParts2(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, int pMillisecond, TimeSpan pOffset);

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified year, month, day, hour, minute, second, millisecond, and offset of a specified calendar.")]

        DateTimeOffset CreateFromDateTimeParts3(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, int pMillisecond, Calendar pCalendar, TimeSpan pOffset);
        //DateTimeOffset CreateFromDateTimeParts3(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, int pMillisecond, GSystem.Globalization.Calendar calendar, TimeSpan pOffset);

        [Description("Initializes a new instance of the DateTimeOffset structure using the specified number of ticks and offset.")]
        DateTimeOffset CreateFromTicks(long pTicks, TimeSpan pOffset);

        // Fields

        DateTimeOffset MaxValue 
        {
            [Description("Represents the greatest possible value of DateTimeOffset. This field is read-only.")]
            get;
        }

        DateTimeOffset MinValue 
        {
            [Description("Represents the earliest possible DateTimeOffset value. This field is read-only.")]
            get;
        }

        // Properties

        DateTimeOffset Now 
        {
            [Description("Gets a DateTimeOffset object that is set to the current date and time on the current computer, with the offset set to the local time's offset from Coordinated Universal Time (UTC).")]
            get;
        }

        DateTimeOffset UtcNow 
        {
            [Description("Gets a DateTimeOffset object whose date and time are set to the current Coordinated Universal Time (UTC) date and time and whose offset is Zero.")]
            get;
        }

        // Methods

        [Description("Compares two DateTimeOffset objects and indicates whether the first is earlier than the second, equal to the second, or later than the second.")]
        int Compare(DateTimeOffset first, DateTimeOffset second);

        [Description("Determines whether two specified DateTimeOffset objects represent the same point in time.")]
        bool Equals(DateTimeOffset first, DateTimeOffset second);

        [Description("Converts the specified Windows file time to an equivalent local time.")]
        DateTimeOffset FromFileTime(long fileTime);

        [Description("Converts a Unix time expressed as the number of milliseconds that have elapsed since 1970-01-01T00:00:00Z to a DateTimeOffset value.")]
        DateTimeOffset FromUnixTimeMilliseconds(long pMilliseconds);

        [Description("Converts a Unix time expressed as the number of seconds that have elapsed since 1970-01-01T00:00:00Z to a DateTimeOffset value.")]
        DateTimeOffset FromUnixTimeSeconds(long pSeconds);

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
        DateTimeOffset ParseExact3(string input, [In] ref string[] formats, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles);

        [Description("Tries to converts a specified string representation of a date and time to its DateTimeOffset equivalent, and returns a value that indicates whether the conversion succeeded.")]
        bool TryParse(string input, out DateTimeOffset result);

        [Description("Tries to convert a specified string representation of a date and time to its DateTimeOffset equivalent, and returns a value that indicates whether the conversion succeeded.")]
        bool TryParse2(string input, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles, out DateTimeOffset result);

        [Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified format, culture-specific format information, and style. The format of the string representation must match the specified format exactly.")]
        bool TryParseExact(string input, string format, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles, out DateTimeOffset result);

        [Description("Converts the specified string representation of a date and time to its DateTimeOffset equivalent using the specified array of formats, culture-specific format information, and style. The format of the string representation must match one of the specified formats exactly.")]
        bool TryParseExact2(string input, [In] ref string[] formats, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles, out DateTimeOffset result);

        // Operatorrs

        [Description("Adds a specified time interval to a DateTimeOffset object that has a specified date and time, and yields a DateTimeOffset object that has new a date and time.")]
        DateTimeOffset Addition (DateTimeOffset dateTimeOffset, TimeSpan timeSpan);

        [Description("Determines whether two specified DateTimeOffset objects represent the same point in time.")]
        bool Equality(DateTimeOffset left, DateTimeOffset right);

        [Description("Determines whether one specified DateTimeOffset object is greater than (or later than) a second specified DateTimeOffset object.")]
        bool GreaterThan(DateTimeOffset left, DateTimeOffset right);

        [Description("Determines whether one specified DateTimeOffset object is greater than or equal to a second specified DateTimeOffset object.")]
        bool GreaterThanOrEqual(DateTimeOffset left, DateTimeOffset right);

        [Description("Defines an implicit conversion of a DateTime object to a DateTimeOffset object.")]
        DateTimeOffset Implicit(DateTime dateTime);

        [Description("Determines whether two specified DateTimeOffset objects refer to different points in time.")]
        bool Inequality(DateTimeOffset left, DateTimeOffset right);

        [Description("Determines whether one specified DateTimeOffset object is less than a second specified DateTimeOffset object.")]
        bool LessThan(DateTimeOffset left, DateTimeOffset right);

        [Description("Determines whether one specified DateTimeOffset object is less than or equal to a second specified DateTimeOffset object.")]
        bool LessThanOrEqual(DateTimeOffset left, DateTimeOffset right);

        [Description("Subtracts one DateTimeOffset object from another and yields a time interval.")]
        TimeSpan Subtraction(DateTimeOffset left, DateTimeOffset right);

        [Description("Subtracts a specified time interval from a specified date and time, and yields a new date and time.")]
        DateTimeOffset Subtraction2(DateTimeOffset dateTimeOffset, TimeSpan timeSpan);
    }
}
