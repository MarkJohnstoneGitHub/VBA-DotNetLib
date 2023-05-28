using GSystem = global::System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System;

namespace DotNetLib.System
{
    // https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset?view=netframework-4.8.1
    // https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/DateTimeOffset.cs

    // TODO : Explict Interface Implementations

    [ComVisible(true)]
    [Description("Represents a point in time, typically expressed as a date and time of day, relative to Coordinated Universal Time (UTC).")]
    [Guid("27660912-6101-4779-B4E0-24F2B164B334")]
    [ProgId("DotNetLib.System.DateTimeOffset")]
    [ClassInterface(ClassInterfaceType.None)]
    public class DateTimeOffset : IDateTimeOffset
    {
        private GSystem.DateTimeOffset dateTimeOffsetObject;

        public static DateTimeOffset dtoMaxValue = new DateTimeOffset(GSystem.DateTimeOffset.MaxValue);
        public static readonly DateTimeOffset dtoMinValue = new DateTimeOffset(GSystem.DateTimeOffset.MinValue);

        // Constructors

        internal DateTimeOffset(GSystem.DateTimeOffset dateTimeOffsetObject)
        {
            this.dateTimeOffsetObject = dateTimeOffsetObject;
        }

        public DateTimeOffset()
        {
            dateTimeOffsetObject = new GSystem.DateTimeOffset();
        }

        public DateTimeOffset(DateTime dateTime)
        {
            dateTimeOffsetObject = new GSystem.DateTimeOffset(dateTime.DateTimeObject);
        }

        public DateTimeOffset(DateTime dateTime, TimeSpan offset)
        {
            dateTimeOffsetObject = new GSystem.DateTimeOffset(dateTime.DateTimeObject, offset.TimeSpanObject);
        }

        public DateTimeOffset(long ticks, TimeSpan offset)
        {
            dateTimeOffsetObject = new GSystem.DateTimeOffset(ticks, offset.TimeSpanObject);
        }

        public DateTimeOffset(int year, int month, int day, int hour, int minute, int second, TimeSpan offset)
        {
            dateTimeOffsetObject = new GSystem.DateTimeOffset(year, month, day, hour, minute, second, offset.TimeSpanObject);
        }

        public DateTimeOffset(int year, int month, int day, int hour, int minute, int second, int millisecond, TimeSpan offset)
        {
            dateTimeOffsetObject = new GSystem.DateTimeOffset(year, month, day, hour, minute, second, millisecond, offset.TimeSpanObject);
        }

        public DateTimeOffset(int year, int month, int day, int hour, int minute, int second, int millisecond, GSystem.Globalization.Calendar calendar, TimeSpan offset)
        {
            dateTimeOffsetObject = new GSystem.DateTimeOffset(year, month, day, hour, minute, second, millisecond, calendar, offset.TimeSpanObject);
        }

        public DateTimeOffset CreateFromDateTime(DateTime dateTime)
        {
            return new DateTimeOffset(dateTime);
        }

        public DateTimeOffset CreateFromDateTime2(DateTime dateTime, TimeSpan offset)
        {
            return new DateTimeOffset(dateTime, offset);
        }

        public DateTimeOffset CreateFromDateTimeParts(int year, int month, int day, int hour, int minute, int second, TimeSpan offset)
        {
            return new DateTimeOffset(year, month, day, hour, minute, second, offset);
        }

        public DateTimeOffset CreateFromDateTimeParts2(int year, int month, int day, int hour, int minute, int second, int millisecond, TimeSpan offset)
        {
            return new DateTimeOffset(year, month, day, hour, minute, second, millisecond, offset);
        }

        public DateTimeOffset CreateFromDateTimeParts3(int year, int month, int day, int hour, int minute, int second, int millisecond, GSystem.Globalization.Calendar calendar, TimeSpan offset)
        {
            return new DateTimeOffset(year, month, day, hour, minute, second, millisecond, calendar, offset);
        }

        public DateTimeOffset CreateFromTicks(long ticks, TimeSpan offset)
        {
            return new DateTimeOffset(ticks, offset);
        }

        // Fields
        public DateTimeOffset MaxValue => dtoMaxValue;
        public DateTimeOffset MinValue => dtoMinValue;

        // Properties

        internal GSystem.DateTimeOffset DateTimeOffsetObject
        {
            get { return this.dateTimeOffsetObject; }
        }

        public DateTime Date
        {
            get { return new DateTime(this.dateTimeOffsetObject.Date); }
        }

        public DateTime DateTime
        {
            get { return new DateTime(this.dateTimeOffsetObject.DateTime); }
        }

        public int Day => this.dateTimeOffsetObject.Day;

        public DayOfWeek DayOfWeek => (DayOfWeek)this.dateTimeOffsetObject.DayOfWeek;


        public int DayOfYear => this.dateTimeOffsetObject.DayOfYear;

        public int Hour => this.dateTimeOffsetObject.Hour;

        public DateTime LocalDateTime
        {
            get { return new DateTime(this.dateTimeOffsetObject.LocalDateTime); }
        }

        public int Millisecond => this.dateTimeOffsetObject.Millisecond;

        public int Minute => this.dateTimeOffsetObject.Minute;

        public int Month => this.dateTimeOffsetObject.Month;

        public DateTimeOffset Now
        {
            get { return new DateTimeOffset(GSystem.DateTimeOffset.Now); }
        }

        public TimeSpan Offset
        {
            get { return new TimeSpan(this.dateTimeOffsetObject.Offset); }
        }

        public int Second => this.dateTimeOffsetObject.Second;

        public long Ticks => this.dateTimeOffsetObject.Ticks;

        public TimeSpan TimeOfDay
        {
            get { return new TimeSpan(this.dateTimeOffsetObject.TimeOfDay); }
        }

        public DateTime UtcDateTime
        {
            get { return new DateTime(this.dateTimeOffsetObject.UtcDateTime); }
        }

        public DateTimeOffset UtcNow
        {
            get { return new DateTimeOffset(GSystem.DateTimeOffset.UtcNow); }
        }

        public long UtcTicks => this.dateTimeOffsetObject.UtcTicks;

        public int Year => this.dateTimeOffsetObject.Year;

        // Methods

        public DateTimeOffset Add(TimeSpan timeSpan)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.Add(timeSpan.TimeSpanObject));
        }

        public DateTimeOffset AddDays(double days)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.AddDays(days));
        }

        public DateTimeOffset AddMilliseconds(double milliseconds)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.AddMilliseconds(milliseconds));
        }

        public DateTimeOffset AddMinutes(double minutes)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.AddMinutes(minutes));
        }

        public DateTimeOffset AddMonths(int months)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.AddMonths(months));
        }

        public DateTimeOffset AddSeconds(double seconds)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.AddSeconds(seconds));
        }

        public DateTimeOffset AddTicks(long ticks)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.AddTicks(ticks));
        }

        public DateTimeOffset AddYears(int years)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.AddYears(years));
        }

        public int Compare(DateTimeOffset first, DateTimeOffset second)
        {
            return GSystem.DateTimeOffset.Compare(first.DateTimeOffsetObject, second.DateTimeOffsetObject);
        }

        public int CompareTo(DateTimeOffset other)
        {
            return this.dateTimeOffsetObject.CompareTo(other.dateTimeOffsetObject);
        }

        public bool Equals(DateTimeOffset other)
        {
            return this.dateTimeOffsetObject.Equals(other.dateTimeOffsetObject);
        }

        // Checks if this DateTimeOffset is equal to a given object. Returns
        // true if the given object is a boxed DateTimeOffset and its value
        // is equal to the value of this DateTimeOffset. Returns false
        // otherwise.
        //
        // TODO : Check implementation
        public bool Equals2(object obj)
        {
            return obj is DateTimeOffset && UtcDateTime.Equals(((DateTimeOffset)obj).UtcDateTime);
        }

        public bool Equals3(DateTimeOffset first, DateTimeOffset second)
        {
            return GSystem.DateTimeOffset.Equals(first.DateTimeOffsetObject, second.DateTimeOffsetObject);
        }

        public bool EqualsExact(DateTimeOffset other)
        {
            return this.dateTimeOffsetObject.EqualsExact(other.dateTimeOffsetObject);
        }

        public DateTimeOffset FromFileTime(long fileTime)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.FromFileTime(fileTime));
        }

        public DateTimeOffset FromUnixTimeMilliseconds(long milliseconds)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.FromUnixTimeMilliseconds(milliseconds));
        }

        public DateTimeOffset FromUnixTimeSeconds(long seconds)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.FromUnixTimeSeconds(seconds));
        }

        public override int GetHashCode()
        {
            return this.dateTimeOffsetObject.GetHashCode();
        }

        public DateTimeOffset Parse(string input)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.Parse(input));
        }

        public DateTimeOffset Parse2(string input, IFormatProvider formatProvider)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.Parse(input, formatProvider));
        }

        public DateTimeOffset Parse3(string input, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.Parse(input, formatProvider, styles));
        }

        public DateTimeOffset ParseExact(string input, string format, IFormatProvider formatProvider)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.ParseExact(input, format, formatProvider));
        }

        public DateTimeOffset ParseExact2(string input, string format, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.ParseExact(input, format, formatProvider, styles));
        }

        public DateTimeOffset ParseExact3(string input, string[] formats, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.ParseExact(input, formats, formatProvider, styles));
        }

        public TimeSpan Subtract(DateTimeOffset value)
        {
            return new TimeSpan(this.dateTimeOffsetObject.Subtract(value.dateTimeOffsetObject));
        }

        public DateTimeOffset Subtract2(TimeSpan value)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.Subtract(value.TimeSpanObject));
        }

        public long ToFileTime()
        {
            return this.dateTimeOffsetObject.ToFileTime();
        }

        public DateTimeOffset ToLocalTime()
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.ToLocalTime());
        }

        public DateTimeOffset ToOffset(TimeSpan offset)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.ToOffset(offset.TimeSpanObject));
        }

        public override string ToString()
        {
            return this.dateTimeOffsetObject.ToString();
        }

        public string ToString2(string format)
        {
            return this.dateTimeOffsetObject.ToString(format);
        }

        public string ToString3(IFormatProvider formatProvider)
        {
            return this.dateTimeOffsetObject.ToString(formatProvider);
        }

        public string ToString4(string format, IFormatProvider formatProvider)
        {
            return this.dateTimeOffsetObject.ToString(format, formatProvider);
        }

        public DateTimeOffset ToUniversalTime()
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.ToUniversalTime());
        }

        public long ToUnixTimeMilliseconds()
        {
            return this.dateTimeOffsetObject.ToUnixTimeMilliseconds();
        }

        public long ToUnixTimeSeconds()
        {
            return this.dateTimeOffsetObject.ToUnixTimeSeconds();
        }

        // TODO: public bool TryFormat(Span<char> destination, out int charsWritten, ReadOnlySpan<char> format = default, IFormatProvider? formatProvider = default);

        public bool TryParse(string input, out DateTimeOffset result)
        {
            bool pvtTryParse = GSystem.DateTimeOffset.TryParse(input, out GSystem.DateTimeOffset outResult);
            result = new DateTimeOffset(outResult);
            return pvtTryParse;
        }

        public bool TryParse2(string input, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles, out DateTimeOffset result)
        {
            bool pvtTryParse = GSystem.DateTimeOffset.TryParse(input, formatProvider, styles,out GSystem.DateTimeOffset outResult);
            result = new DateTimeOffset(outResult);
            return pvtTryParse;
        }

        public bool TryParseExact(string input, string format, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles, out DateTimeOffset result)
        {
            bool pvtTryParseExact = GSystem.DateTimeOffset.TryParseExact(input, format, formatProvider, styles, out GSystem.DateTimeOffset outResult);
            result = new DateTimeOffset(outResult);
            return pvtTryParseExact;
        }

        public bool TryParseExact2(string input, string[] formats, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles, out DateTimeOffset result)
        {
            bool pvtTryParseExact = GSystem.DateTimeOffset.TryParseExact(input, formats, formatProvider, styles, out GSystem.DateTimeOffset outResult);
            result = new DateTimeOffset(outResult);
            return pvtTryParseExact;
        }

        // Operators

        public DateTimeOffset Addition(DateTimeOffset dateTimeOffset, TimeSpan timeSpan)
        {
            return new DateTimeOffset(dateTimeOffset.DateTimeOffsetObject + timeSpan.TimeSpanObject);
        }

        public bool Equality(DateTimeOffset left, DateTimeOffset right)
        {
            return (left.DateTimeOffsetObject == right.DateTimeOffsetObject);
        }

        public bool GreaterThan(DateTimeOffset left, DateTimeOffset right)
        {
            return (left.DateTimeOffsetObject > right.DateTimeOffsetObject);
        }

        public bool GreaterThanOrEqual(DateTimeOffset left, DateTimeOffset right)
        {
            return (left.DateTimeOffsetObject  >= right.DateTimeOffsetObject);
        }

        public DateTimeOffset Implicit(DateTime dateTime) =>
            new DateTimeOffset(dateTime);

        public bool Inequality(DateTimeOffset left, DateTimeOffset right)
        {
            return (left.DateTimeOffsetObject != right.DateTimeOffsetObject);
        }

        public bool LessThan(DateTimeOffset left, DateTimeOffset right)
        {
            return (left.DateTimeOffsetObject < right.DateTimeOffsetObject);
        }

        public bool LessThanOrEqual(DateTimeOffset left, DateTimeOffset right)
        {
            return (left.DateTimeOffsetObject <= right.DateTimeOffsetObject);
        }

        public TimeSpan Subtraction(DateTimeOffset left, DateTimeOffset right)
        {
            return new TimeSpan(left.DateTimeOffsetObject - right.DateTimeOffsetObject);
        }

        public DateTimeOffset Subtraction2(DateTimeOffset dateTimeOffset, TimeSpan timeSpan)
        {
            return new DateTimeOffset(dateTimeOffset.DateTimeOffsetObject - timeSpan.TimeSpanObject);
        }


    }


}
