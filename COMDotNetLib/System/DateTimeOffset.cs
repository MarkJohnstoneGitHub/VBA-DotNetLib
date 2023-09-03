// https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset?view=netframework-4.8.1
// https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/DateTimeOffset.cs

using GSystem = global::System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System;
using DotNetLib.System.Globalization;

namespace DotNetLib.System
{
    // TODO : Explict Interface Implementations

    [ComVisible(true)]
    [Description("Represents a point in time, typically expressed as a date and time of day, relative to Coordinated Universal Time (UTC).")]
    [Guid("27660912-6101-4779-B4E0-24F2B164B334")]
    [ProgId("DotNetLib.System.DateTimeOffset")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDateTimeOffset))]
    public class DateTimeOffset : IDateTimeOffset
    {
        private GSystem.DateTimeOffset _dateTimeOffset;

        // Static Fields
        private static readonly DateTimeOffset dtoMaxValue = new DateTimeOffset(GSystem.DateTimeOffset.MaxValue);
        private static readonly DateTimeOffset dtoMinValue = new DateTimeOffset(GSystem.DateTimeOffset.MinValue);

        // Constructors

        internal DateTimeOffset(GSystem.DateTimeOffset dateTimeOffsetObject)
        {
            this._dateTimeOffset = dateTimeOffsetObject;
        }

        public DateTimeOffset()
        {
            _dateTimeOffset = new GSystem.DateTimeOffset();
        }

        public DateTimeOffset(DateTime dateTime)
        {
            _dateTimeOffset = new GSystem.DateTimeOffset(dateTime.DateTimeObject);
        }

        public DateTimeOffset(DateTime dateTime, TimeSpan offset)
        {
            _dateTimeOffset = new GSystem.DateTimeOffset(dateTime.DateTimeObject, offset.TimeSpanObject);
        }

        public DateTimeOffset(long ticks, TimeSpan offset)
        {
            _dateTimeOffset = new GSystem.DateTimeOffset(ticks, offset.TimeSpanObject);
        }

        public DateTimeOffset(int year, int month, int day, int hour, int minute, int second, TimeSpan offset)
        {
            _dateTimeOffset = new GSystem.DateTimeOffset(year, month, day, hour, minute, second, offset.TimeSpanObject);
        }

        public DateTimeOffset(int year, int month, int day, int hour, int minute, int second, int millisecond, TimeSpan offset)
        {
            _dateTimeOffset = new GSystem.DateTimeOffset(year, month, day, hour, minute, second, millisecond, offset.TimeSpanObject);
        }

        public DateTimeOffset(int year, int month, int day, int hour, int minute, int second, int millisecond, GSystem.Globalization.Calendar calendar, TimeSpan offset)
        {
            _dateTimeOffset = new GSystem.DateTimeOffset(year, month, day, hour, minute, second, millisecond, calendar, offset.TimeSpanObject);
        }

        // Fields
        public static DateTimeOffset MaxValue => dtoMaxValue;
        public static DateTimeOffset MinValue => dtoMinValue;

        // Properties

        internal GSystem.DateTimeOffset DateTimeOffsetObject
        {
            get { return _dateTimeOffset; }
        }

        //Note: Renamed to Date to avoid issues with VBA reserved word Date
        public DateTime Date            
        {
            get { return new DateTime(_dateTimeOffset.Date); }
        }

        public DateTime DateTime
        {
            get { return new DateTime(_dateTimeOffset.DateTime); }
        }

        public int Day => _dateTimeOffset.Day;

        public DayOfWeek DayOfWeek => (DayOfWeek)_dateTimeOffset.DayOfWeek;


        public int DayOfYear => _dateTimeOffset.DayOfYear;

        public int Hour => _dateTimeOffset.Hour;

        public DateTime LocalDateTime
        {
            get { return new DateTime(_dateTimeOffset.LocalDateTime); }
        }

        public int Millisecond => _dateTimeOffset.Millisecond;

        public int Minute => _dateTimeOffset.Minute;

        public int Month => _dateTimeOffset.Month;

        public static DateTimeOffset Now
        {
            get { return new DateTimeOffset(GSystem.DateTimeOffset.Now); }
        }

        public TimeSpan Offset
        {
            get { return new TimeSpan(_dateTimeOffset.Offset); }
        }

        public int Second => _dateTimeOffset.Second;

        public long Ticks => _dateTimeOffset.Ticks;

        public TimeSpan TimeOfDay
        {
            get { return new TimeSpan(_dateTimeOffset.TimeOfDay); }
        }

        public DateTime UtcDateTime
        {
            get { return new DateTime(_dateTimeOffset.UtcDateTime); }
        }

        public static DateTimeOffset UtcNow
        {
            get { return new DateTimeOffset(GSystem.DateTimeOffset.UtcNow); }
        }

        public long UtcTicks => _dateTimeOffset.UtcTicks;

        public int Year => _dateTimeOffset.Year;

        // Methods

        public DateTimeOffset Add(TimeSpan ts)
        {
            return new DateTimeOffset(_dateTimeOffset.Add(ts.TimeSpanObject));
        }

        public DateTimeOffset AddDays(double days)
        {
            return new DateTimeOffset(_dateTimeOffset.AddDays(days));
        }

        public DateTimeOffset AddHours(double hours)
        {
            return new DateTimeOffset(_dateTimeOffset.AddHours(hours));
        }

        public DateTimeOffset AddMilliseconds(double pMilliseconds)
        {
            return new DateTimeOffset(_dateTimeOffset.AddMilliseconds(pMilliseconds));
        }

        public DateTimeOffset AddMinutes(double minutes)
        {
            return new DateTimeOffset(_dateTimeOffset.AddMinutes(minutes));
        }

        public DateTimeOffset AddMonths(int months)
        {
            return new DateTimeOffset(_dateTimeOffset.AddMonths(months));
        }

        public DateTimeOffset AddSeconds(double pSeconds)
        {
            return new DateTimeOffset(_dateTimeOffset.AddSeconds(pSeconds));
        }

        public DateTimeOffset AddTicks(long ticks)
        {
            return new DateTimeOffset(_dateTimeOffset.AddTicks(ticks));
        }

        public DateTimeOffset AddYears(int years)
        {
            return new DateTimeOffset(_dateTimeOffset.AddYears(years));
        }

        public static int Compare(DateTimeOffset first, DateTimeOffset second)
        {
            return GSystem.DateTimeOffset.Compare(first.DateTimeOffsetObject, second.DateTimeOffsetObject);
        }

        public int CompareTo(DateTimeOffset other)
        {
            return _dateTimeOffset.CompareTo(other.DateTimeOffsetObject);
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
            return _dateTimeOffset.Equals(other.DateTimeOffsetObject);
        }

        // Checks if this DateTimeOffset is equal to a given object. Returns
        // true if the given object is a boxed DateTimeOffset and its value
        // is equal to the value of this DateTimeOffset. Returns false
        // otherwise.
        //
        public bool Equals2(object obj)
        {
            return obj is DateTimeOffset dto && _dateTimeOffset == dto.DateTimeOffsetObject;
        }

        public static bool Equals(DateTimeOffset first, DateTimeOffset second)
        {
            return GSystem.DateTimeOffset.Equals(first.DateTimeOffsetObject, second.DateTimeOffsetObject);
        }

        public bool EqualsExact(DateTimeOffset other)
        {
            return _dateTimeOffset.EqualsExact(other.DateTimeOffsetObject);
        }

        public static DateTimeOffset FromFileTime(long fileTime)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.FromFileTime(fileTime));
        }

        public static DateTimeOffset FromUnixTimeMilliseconds(long pMilliseconds)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.FromUnixTimeMilliseconds(pMilliseconds));
        }

        public static DateTimeOffset FromUnixTimeSeconds(long pSeconds)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.FromUnixTimeSeconds(pSeconds));
        }

        public override int GetHashCode()
        {
            return _dateTimeOffset.GetHashCode();
        }

        public static DateTimeOffset Parse(string input)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.Parse(input));
        }

        public static DateTimeOffset Parse(string input, IFormatProvider formatProvider)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.Parse(input, DateTimeFormatInfo.GetFormatProvider(formatProvider)));
        }

        public static DateTimeOffset Parse(string input, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.Parse(input, DateTimeFormatInfo.GetFormatProvider(formatProvider), styles));
        }

        public static DateTimeOffset ParseExact(string input, string format, IFormatProvider formatProvider)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.ParseExact(input, format, DateTimeFormatInfo.GetFormatProvider(formatProvider)));
        }

        public static DateTimeOffset ParseExact(string input, string format, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.ParseExact(input, format, DateTimeFormatInfo.GetFormatProvider(formatProvider), styles));
        }

        public static DateTimeOffset ParseExact(string input, [In] ref string[] formats, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.ParseExact(input, formats, DateTimeFormatInfo.GetFormatProvider(formatProvider), styles));
        }

        public TimeSpan Subtract(DateTimeOffset value)
        {
            return new TimeSpan(_dateTimeOffset.Subtract(value.DateTimeOffsetObject));
        }

        public DateTimeOffset Subtract2(TimeSpan value)
        {
            return new DateTimeOffset(_dateTimeOffset.Subtract(value.TimeSpanObject));
        }

        public long ToFileTime()
        {
            return _dateTimeOffset.ToFileTime();
        }

        public DateTimeOffset ToLocalTime()
        {
            return new DateTimeOffset(_dateTimeOffset.ToLocalTime());
        }

        public DateTimeOffset ToOffset(TimeSpan offset)
        {
            return new DateTimeOffset(_dateTimeOffset.ToOffset(offset.TimeSpanObject));
        }

        public override string ToString()
        {
            return _dateTimeOffset.ToString();
        }

        public string ToString2(string format)
        {
            return _dateTimeOffset.ToString(format);
        }

        public string ToString3(IFormatProvider formatProvider)
        {
            return _dateTimeOffset.ToString(DateTimeFormatInfo.GetFormatProvider(formatProvider));
        }

        public string ToString4(string format, IFormatProvider formatProvider)
        {
            return _dateTimeOffset.ToString(format, DateTimeFormatInfo.GetFormatProvider(formatProvider));
        }

        public DateTimeOffset ToUniversalTime()
        {
            return new DateTimeOffset(_dateTimeOffset.ToUniversalTime());
        }

        public long ToUnixTimeMilliseconds()
        {
            return _dateTimeOffset.ToUnixTimeMilliseconds();
        }

        public long ToUnixTimeSeconds()
        {
            return _dateTimeOffset.ToUnixTimeSeconds();
        }

        // TODO: public bool TryFormat(Span<char> destination, out int charsWritten, ReadOnlySpan<char> format = default, IFormatProvider? formatProvider = default);

        public static bool TryParse(string input, out DateTimeOffset result)
        {
            bool pvtTryParse = GSystem.DateTimeOffset.TryParse(input, out GSystem.DateTimeOffset outResult);
            result = new DateTimeOffset(outResult);
            return pvtTryParse;
        }

        public static bool TryParse(string input, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles, out DateTimeOffset result)
        {
            bool pvtTryParse = GSystem.DateTimeOffset.TryParse(input, DateTimeFormatInfo.GetFormatProvider(formatProvider), styles,out GSystem.DateTimeOffset outResult);
            result = new DateTimeOffset(outResult);
            return pvtTryParse;
        }

        public static bool TryParseExact(string input, string format, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles, out DateTimeOffset result)
        {
            bool pvtTryParseExact = GSystem.DateTimeOffset.TryParseExact(input, format, DateTimeFormatInfo.GetFormatProvider(formatProvider), styles, out GSystem.DateTimeOffset outResult);
            result = new DateTimeOffset(outResult);
            return pvtTryParseExact;
        }

        public static bool TryParseExact(string input, string[] formats, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles, out DateTimeOffset result)
        {
            bool pvtTryParseExact = GSystem.DateTimeOffset.TryParseExact(input, formats, DateTimeFormatInfo.GetFormatProvider(formatProvider), styles, out GSystem.DateTimeOffset outResult);
            result = new DateTimeOffset(outResult);
            return pvtTryParseExact;
        }

        // Operators

        public static DateTimeOffset Addition(DateTimeOffset dateTimeOffset, TimeSpan timeSpan)
        {
            return new DateTimeOffset(dateTimeOffset.DateTimeOffsetObject + timeSpan.TimeSpanObject);
        }

        public static bool Equality(DateTimeOffset left, DateTimeOffset right)
        {
            return (left.DateTimeOffsetObject == right.DateTimeOffsetObject);
        }

        public static bool GreaterThan(DateTimeOffset left, DateTimeOffset right)
        {
            return (left.DateTimeOffsetObject > right.DateTimeOffsetObject);
        }

        public static bool GreaterThanOrEqual(DateTimeOffset left, DateTimeOffset right)
        {
            return (left.DateTimeOffsetObject  >= right.DateTimeOffsetObject);
        }

        public static DateTimeOffset Implicit(DateTime dateTime) => new DateTimeOffset(dateTime);

        public static bool Inequality(DateTimeOffset left, DateTimeOffset right)
        {
            return (left.DateTimeOffsetObject != right.DateTimeOffsetObject);
        }

        public static bool LessThan(DateTimeOffset left, DateTimeOffset right)
        {
            return (left.DateTimeOffsetObject < right.DateTimeOffsetObject);
        }

        public static bool LessThanOrEqual(DateTimeOffset left, DateTimeOffset right)
        {
            return (left.DateTimeOffsetObject <= right.DateTimeOffsetObject);
        }

        public static TimeSpan Subtraction(DateTimeOffset left, DateTimeOffset right)
        {
            return new TimeSpan(left.DateTimeOffsetObject - right.DateTimeOffsetObject);
        }

        public static DateTimeOffset Subtraction(DateTimeOffset dateTimeOffset, TimeSpan timeSpan)
        {
            return new DateTimeOffset(dateTimeOffset.DateTimeOffsetObject - timeSpan.TimeSpanObject);
        }
    }
}
