// https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset?view=netframework-4.8.1
// https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/DateTimeOffset.cs

using GSystem = global::System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System;

namespace DotNetLib.System
{
    // TODO : Explict Interface Implementations

    [ComVisible(true)]
    [Description("Represents a point in time, typically expressed as a date and time of day, relative to Coordinated Universal Time (UTC).")]
    [Guid("27660912-6101-4779-B4E0-24F2B164B334")]
    [ProgId("DotNetLib.System.DateTimeOffset")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDateTimeOffset))]
    public class DateTimeOffset : IDateTimeOffset, IDateTimeOffsetSingleton
    {
        private GSystem.DateTimeOffset dateTimeOffsetObject;

        // Static Fields
        private static readonly IDateTimeOffset dtoMaxValue = new DateTimeOffset(GSystem.DateTimeOffset.MaxValue);
        private static readonly IDateTimeOffset dtoMinValue = new DateTimeOffset(GSystem.DateTimeOffset.MinValue);

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

        public IDateTimeOffset CreateFromDateTime(DateTime pDateTime)
        {
            return new DateTimeOffset(pDateTime);
        }

        public IDateTimeOffset CreateFromDateTime2(DateTime pDateTime, TimeSpan pOffset)
        {
            return new DateTimeOffset(pDateTime, pOffset);
        }

        public IDateTimeOffset CreateFromDateTimeParts(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, TimeSpan pOffset)
        {
            return new DateTimeOffset(pYear, pMonth, pDay, pHour, pMinute, pSecond, pOffset);
        }

        public IDateTimeOffset CreateFromDateTimeParts2(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, int pMillisecond, TimeSpan pOffset)
        {
            return new DateTimeOffset(pYear, pMonth, pDay, pHour, pMinute, pSecond, pMillisecond, pOffset);
        }

        public IDateTimeOffset CreateFromDateTimeParts3(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, int pMillisecond, GSystem.Globalization.Calendar pCalendar, TimeSpan pOffset)
        {
            return new DateTimeOffset(pYear, pMonth, pDay, pHour, pMinute, pSecond, pMillisecond, pCalendar, pOffset);
        }

        public IDateTimeOffset CreateFromTicks(long pTicks, TimeSpan pOffset)
        {
            return new DateTimeOffset(pTicks, pOffset);
        }

        // Fields
        public IDateTimeOffset MaxValue => dtoMaxValue;
        public IDateTimeOffset MinValue => dtoMinValue;

        // Properties

        internal GSystem.DateTimeOffset DateTimeOffsetObject
        {
            get { return this.dateTimeOffsetObject; }
        }

        //Note: Renamed to DateOnly to avoid issues with VBA reserved word Date
        public IDateTime DateOnly            
        {
            get { return new DateTime(this.dateTimeOffsetObject.Date); }
        }

        public IDateTime DateTime
        {
            get { return new DateTime(this.dateTimeOffsetObject.DateTime); }
        }

        public int Day => this.dateTimeOffsetObject.Day;

        public DayOfWeek DayOfWeek => (DayOfWeek)this.dateTimeOffsetObject.DayOfWeek;


        public int DayOfYear => this.dateTimeOffsetObject.DayOfYear;

        public int Hour => this.dateTimeOffsetObject.Hour;

        public IDateTime LocalDateTime
        {
            get { return new DateTime(this.dateTimeOffsetObject.LocalDateTime); }
        }

        public int Millisecond => this.dateTimeOffsetObject.Millisecond;

        public int Minute => this.dateTimeOffsetObject.Minute;

        public int Month => this.dateTimeOffsetObject.Month;

        public IDateTimeOffset Now
        {
            get { return new DateTimeOffset(GSystem.DateTimeOffset.Now); }
        }

        public ITimeSpan Offset
        {
            get { return new TimeSpan(this.dateTimeOffsetObject.Offset); }
        }

        public int Second => this.dateTimeOffsetObject.Second;

        public long Ticks => this.dateTimeOffsetObject.Ticks;

        public ITimeSpan TimeOfDay
        {
            get { return new TimeSpan(this.dateTimeOffsetObject.TimeOfDay); }
        }

        public IDateTime UtcDateTime
        {
            get { return new DateTime(this.dateTimeOffsetObject.UtcDateTime); }
        }

        public IDateTimeOffset UtcNow
        {
            get { return new DateTimeOffset(GSystem.DateTimeOffset.UtcNow); }
        }

        public long UtcTicks => this.dateTimeOffsetObject.UtcTicks;

        public int Year => this.dateTimeOffsetObject.Year;

        // Methods

        public IDateTimeOffset Add(TimeSpan ts)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.Add(ts.TimeSpanObject));
        }

        public IDateTimeOffset AddDays(double days)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.AddDays(days));
        }

        public IDateTimeOffset AddHours(double hours)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.AddHours(hours));
        }

        public IDateTimeOffset AddMilliseconds(double pMilliseconds)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.AddMilliseconds(pMilliseconds));
        }

        public IDateTimeOffset AddMinutes(double minutes)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.AddMinutes(minutes));
        }

        public IDateTimeOffset AddMonths(int months)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.AddMonths(months));
        }

        public IDateTimeOffset AddSeconds(double pSeconds)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.AddSeconds(pSeconds));
        }

        public IDateTimeOffset AddTicks(long ticks)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.AddTicks(ticks));
        }

        public IDateTimeOffset AddYears(int years)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.AddYears(years));
        }

        public int Compare(DateTimeOffset first, DateTimeOffset second)
        {
            return GSystem.DateTimeOffset.Compare(first.DateTimeOffsetObject, second.DateTimeOffsetObject);
        }

        public int CompareTo(DateTimeOffset other)
        {
            return this.dateTimeOffsetObject.CompareTo(other.DateTimeOffsetObject);
        }

        public int CompareTo2(object value)
        {
            const string Arg_MustBeDateTimeOffset = "Object must be of type (DateTimeOffset.";

            if (value == null) return 1;
            if (!(value is DateTimeOffset dto))
            {
                throw new ArgumentException(Arg_MustBeDateTimeOffset);
            }
            return this.CompareTo(dto);
        }

        public bool Equals(DateTimeOffset other)
        {
            return this.dateTimeOffsetObject.Equals(other.DateTimeOffsetObject);
        }

        // Checks if this DateTimeOffset is equal to a given object. Returns
        // true if the given object is a boxed DateTimeOffset and its value
        // is equal to the value of this DateTimeOffset. Returns false
        // otherwise.
        //
        public bool Equals2(object obj)
        {
            return obj is DateTimeOffset dto && this.dateTimeOffsetObject == dto.DateTimeOffsetObject;
        }

        public bool Equals3(DateTimeOffset first, DateTimeOffset second)
        {
            return GSystem.DateTimeOffset.Equals(first.DateTimeOffsetObject, second.DateTimeOffsetObject);
        }

        public bool EqualsExact(DateTimeOffset other)
        {
            return this.dateTimeOffsetObject.EqualsExact(other.DateTimeOffsetObject);
        }

        public IDateTimeOffset FromFileTime(long fileTime)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.FromFileTime(fileTime));
        }

        public IDateTimeOffset FromUnixTimeMilliseconds(long pMilliseconds)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.FromUnixTimeMilliseconds(pMilliseconds));
        }

        public IDateTimeOffset FromUnixTimeSeconds(long pSeconds)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.FromUnixTimeSeconds(pSeconds));
        }

        public override int GetHashCode()
        {
            return this.dateTimeOffsetObject.GetHashCode();
        }

        public IDateTimeOffset Parse(string input)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.Parse(input));
        }

        public IDateTimeOffset Parse2(string input, IFormatProvider formatProvider)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.Parse(input, formatProvider));
        }

        public IDateTimeOffset Parse3(string input, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.Parse(input, formatProvider, styles));
        }

        public IDateTimeOffset ParseExact(string input, string format, IFormatProvider formatProvider)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.ParseExact(input, format, formatProvider));
        }

        public IDateTimeOffset ParseExact2(string input, string format, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.ParseExact(input, format, formatProvider, styles));
        }

        public IDateTimeOffset ParseExact3(string input, [In] ref string[] formats, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.ParseExact(input, formats, formatProvider, styles));
        }

        public ITimeSpan Subtract(DateTimeOffset value)
        {
            return new TimeSpan(this.dateTimeOffsetObject.Subtract(value.DateTimeOffsetObject));
        }

        public IDateTimeOffset Subtract2(TimeSpan value)
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.Subtract(value.TimeSpanObject));
        }

        public long ToFileTime()
        {
            return this.dateTimeOffsetObject.ToFileTime();
        }

        public IDateTimeOffset ToLocalTime()
        {
            return new DateTimeOffset(this.dateTimeOffsetObject.ToLocalTime());
        }

        public IDateTimeOffset ToOffset(TimeSpan offset)
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

        public IDateTimeOffset ToUniversalTime()
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

        public bool TryParseExact2(string input, [In] ref string[] formats, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles, out DateTimeOffset result)
        {
            bool pvtTryParseExact = GSystem.DateTimeOffset.TryParseExact(input, formats, formatProvider, styles, out GSystem.DateTimeOffset outResult);
            result = new DateTimeOffset(outResult);
            return pvtTryParseExact;
        }

        // Operators

        public IDateTimeOffset Addition(DateTimeOffset dateTimeOffset, TimeSpan timeSpan)
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

        public IDateTimeOffset Implicit(DateTime dateTime) =>
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

        public ITimeSpan Subtraction(DateTimeOffset left, DateTimeOffset right)
        {
            return new TimeSpan(left.DateTimeOffsetObject - right.DateTimeOffsetObject);
        }

        public IDateTimeOffset Subtraction2(DateTimeOffset dateTimeOffset, TimeSpan timeSpan)
        {
            return new DateTimeOffset(dateTimeOffset.DateTimeOffsetObject - timeSpan.TimeSpanObject);
        }
    }
}
