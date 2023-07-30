using System.Runtime.InteropServices;
using GSystem = global::System; // https://stackoverflow.com/questions/5681537/namespace-conflict-in-c-sharp
using System.ComponentModel;
using System;

namespace DotNetLib.System
{
    // https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1
    // https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/DateTime.cs

    // TODO : Explict Interface Implementations

    [ComVisible(true)]
    [Guid("72E3AB6F-9742-4F2F-8FA2-43FEEB1CC788")]
    [ProgId("DotNetLib.System.DateTime")]
    [Description("Represents an instant in time, typically expressed as a date and time of day.")]
    [ClassInterface(ClassInterfaceType.None)]
    public class DateTime : IComDateTime
    {
        private GSystem.DateTime dateTimeObject;
        private static readonly DateTime maxValueObject = new DateTime(GSystem.DateTime.MaxValue);
        private static readonly DateTime minValueObject = new DateTime(GSystem.DateTime.MinValue);
        //public static readonly DateTime dtUnixEpoch = new DateTime(GSystem.DateTime.UnixEpoch);  //Not available in .netframework

        //Constructors
        public DateTime()
        {
            dateTimeObject = new GSystem.DateTime();
        }

        internal DateTime(GSystem.DateTime dateTimeObject)
        {
            this.dateTimeObject = dateTimeObject;
        }

        public DateTime(long ticks)
        {
            this.dateTimeObject = new GSystem.DateTime(ticks);
        }

        public DateTime(long ticks, DateTimeKind kind)
        {
            this.dateTimeObject = new GSystem.DateTime(ticks, (GSystem.DateTimeKind)kind);
        }

        public DateTime(int year, int month, int day)
        {
            this.dateTimeObject = new GSystem.DateTime(year, month, day);
        }

        public DateTime(int year, int month, int day, int hour, int minute, int second)
        {
            this.dateTimeObject = new GSystem.DateTime(year, month, day, hour, minute, second);
        }

        public DateTime(int year, int month, int day, int hour, int minute, int second, int millisecond)
        {
            this.dateTimeObject = new GSystem.DateTime(year, month, day, hour, minute, second, millisecond);
        }

        public DateTime(int year, int month, int day, int hour, int minute, int second, DateTimeKind kind)
        {
            this.dateTimeObject = new GSystem.DateTime(year, month, day, hour, minute, second, (GSystem.DateTimeKind)kind);
        }
        public DateTime(int year, int month, int day, int hour, int minute, int second, int millisecond, DateTimeKind kind)
        {
            this.dateTimeObject = new GSystem.DateTime(year, month, day, hour, minute, second, millisecond, (GSystem.DateTimeKind)kind);
        }

        /// <summary>
        /// Initializes a new instance of the DateTime structure to a specified number of ticks and to Coordinated Universal Time (UTC) or local time.
        /// </summary>
        /// <param name="ticks">A date and time expressed in the number of 100-nanosecond intervals that have elapsed since January 1, 0001 at 00:00:00.000 in the Gregorian calendar.</param>
        /// <param name="kind">One of the enumeration values that indicates whether ticks specifies a local time, Coordinated Universal Time (UTC), or neither.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException"> 
        /// ticks is less than DateTime.MinValue or greater than DateTime.MaxValue.
        /// </exception>
        /// <exception cref="ArgumentException"> 
        /// kind is not one of the DateTimeKind values.
        /// </exception>
        /// 
        public DateTime CreateFromTicks(long ticks, System.DateTimeKind kind = System.DateTimeKind.Unspecified)
        {
            return new DateTime(ticks, kind);
        }
        public DateTime CreateFromDate(int year, int month, int day)
        {
            return new DateTime(year, month, day);
        }

        /// <summary>
        /// Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and millisecond.
        /// </summary>
        /// <param name="year">The year (1 through 9999).</param>
        /// <param name="month">The month (1 through 12).</param>
        /// <param name="day">The day (1 through the number of days in month).</param>
        /// <param name="hour">The hours (0 through 23).</param>
        /// <param name="minute">The minutes (0 through 59).</param>
        /// <param name="second">The seconds (0 through 59).</param>
        /// <param name="millisecond">The milliseconds (0 through 999).</param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException">
        /// <paramref name="year"/> is less than 1 or greater than 9999.
        ///
        /// -or-
        ///
        /// <paramref name="month"/> is less than 1 or greater than 12.
        ///
        /// -or-
        ///
        /// <paramref name="day"/> is less than 1 or greater than the number of days in <paramref name="month"/>.
        ///
        /// -or-
        ///
        /// <paramref name="hour"/> is less than 0 or greater than 23.
        ///
        /// -or-
        ///
        /// <paramref name="minute"/> is less than 0 or greater than 59.
        ///
        /// -or-
        ///
        /// <paramref name="second"/> is less than 0 or greater than 59.
        ///
        /// -or-
        ///
        /// <paramref name="millisecond"/> is less than 0 or greater than 999.
        /// </exception>        
        public System.DateTime CreateFromDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond = 0)
        {
            return new System.DateTime(year, month, day, hour, minute, second, millisecond);
        }

        public System.DateTime CreateFromDateTimeKind(int year, int month, int day, int hour, int minute, int second, DateTimeKind kind)
        {
            return new System.DateTime(year, month, day, hour, minute, second, kind);
        }

        public System.DateTime CreateFromDateTimeKind2(int year, int month, int day, int hour, int minute, int second, int millisecond, DateTimeKind kind)
        {
            return new System.DateTime(year, month, day, hour, minute, second, millisecond, kind);
        }

        //Fields

        /// <summary>
        /// Represents the largest possible value of DateTime. This field is read-only.
        /// </summary>
        public DateTime MaxValue => maxValueObject;

        /// <summary>
        /// Represents the smallest possible value of DateTime. This field is read-only.
        /// </summary>
        public DateTime MinValue => minValueObject;

        /// <summary>
        /// The value of this constant is equivalent to 00:00:00.0000000 UTC, January 1, 1970, in the Gregorian calendar. UnixEpoch defines the point in time when Unix time is equal to 0.
        /// </summary>
        //public DateTime UnixEpoch => dtUnixEpoch;

        //Properties
        internal GSystem.DateTime DateTimeObject
        {
            get { return this.dateTimeObject; }
            //set { objDateTime = value; }  // set method
        }

        public DateTime DateOnly => new DateTime(this.dateTimeObject.Date);

        /// <summary>
        /// Gets the day of the month represented by this instance.
        /// </summary>
        public int Day => this.dateTimeObject.Day;

        /// <summary>
        /// Gets the day of the week represented by this instance.
        /// </summary>
        public DayOfWeek DayOfWeek => (DayOfWeek)dateTimeObject.DayOfWeek;

        /// <summary>
        /// Gets the day of the year represented by this instance.
        /// </summary>
        public int DayOfYear => this.dateTimeObject.DayOfYear;

        /// <summary>
        /// Gets the hour component of the date represented by this instance.
        /// </summary>
        public int Hour => this.dateTimeObject.Hour;

        /// <summary>
        /// Gets a value that indicates whether the time represented by this instance is based on local time, Coordinated Universal Time (UTC), or neither.
        /// </summary>
        public DateTimeKind Kind => (DateTimeKind)this.dateTimeObject.Kind;

        /// <summary>
        /// The microseconds component, expressed as a value between 0 and 999.
        /// </summary>
        //public int Microsecond => (int)(this.objDateTime.Microsecond);

        /// <summary>
        /// Gets the milliseconds component of the date represented by this instance.
        /// </summary>
        public int Millisecond => (int)(this.dateTimeObject.Millisecond);

        /// <summary>
        /// Gets the minute component of the date represented by this instance.
        /// </summary>
        public int Minute => (int)(this.dateTimeObject.Minute);

        /// <summary>
        /// Gets the month component of the date represented by this instance.
        /// </summary>
        public int Month => this.dateTimeObject.Month;

        /// <summary>
        /// The nanoseconds component, expressed as a value between 0 and 900 (in increments of 100 nanoseconds).
        /// </summary>
        //public int Nanosecond => this.objDateTime.Nanosecond;

        /// <summary>
        /// Gets a DateTime object that is set to the current date and time on this computer, expressed as the local time.
        /// </summary>
        public DateTime Now => new DateTime(GSystem.DateTime.Now);

        /// <summary>
        /// Gets the seconds component of the date represented by this instance.
        /// </summary>
        public int Second => this.dateTimeObject.Second;

        public TimeSpan TimeOfDay => new TimeSpan(this.dateTimeObject.Ticks);


        /// <summary>
        /// Gets the number of ticks that represent the date and time of this instance.
        /// </summary>
        public long Ticks => this.dateTimeObject.Ticks;

        /// <summary>
        /// Gets the current date.
        /// </summary>
        public DateTime Today => new DateTime(GSystem.DateTime.Today);

        /// <summary>
        /// Gets a DateTime object that is set to the current date and time on this computer, expressed as the Coordinated Universal Time (UTC).
        /// </summary>
        public DateTime UtcNow => new DateTime(GSystem.DateTime.UtcNow);

        /// <summary>
        /// Gets the year component of the date represented by this instance.
        /// </summary>
        public int Year => this.dateTimeObject.Year;

        // Methods

        public DateTime Add(TimeSpan value)
        {
            return new DateTime(this.dateTimeObject.Add(value.TimeSpanObject));
        }

        public DateTime AddDays(double value)
        {
            return new DateTime(this.dateTimeObject.AddDays(value));
        }

        /// <summary>
        /// Returns a new DateTime that adds the specified number of hours to the value of this instance.
        /// </summary>
        /// <param name="value">A number of whole and fractional hours. The value parameter can be negative or positive.</param>
        /// <returns>An object whose value is the sum of the date and time represented by this instance and the number of hours represented by value.</returns>
        public DateTime AddHours(double value)
        {
            return new DateTime(this.dateTimeObject.AddHours(value));
        }

        //public DateTimeTools AddMicoseconds(double value)
        //{
        //    return new DateTimeTools(this.objDateTime.AddMicoseconds(value));
        //}

        /// <summary>
        /// Returns a new DateTime that adds the specified number of milliseconds to the value of this instance.
        /// </summary>
        /// <param name="value">A number of whole and fractional milliseconds. The value parameter can be negative or positive. Note that this value is rounded to the nearest integer.</param>
        /// <returns>An object whose value is the sum of the date and time represented by this instance and the number of milliseconds represented by value.</returns>
        public DateTime AddMilliseconds(double value)
        {
            return new DateTime(this.dateTimeObject.AddMilliseconds(value));
        }

        /// <summary>
        /// Returns a new DateTime that adds the specified number of minutes to the value of this instance.
        /// </summary>
        /// <param name="value">A number of whole and fractional minutes. The value parameter can be negative or positive.</param>
        /// <returns>An object whose value is the sum of the date and time represented by this instance and the number of minutes represented by value.An object whose value is the sum of the date and time represented by this instance and the number of minutes represented by value.</returns>
        public DateTime AddMinutes(double value)
        {
            return new DateTime(this.dateTimeObject.AddMinutes(value));
        }

        /// <summary>
        /// Returns a new DateTime that adds the specified number of months to the value of this instance.
        /// </summary>
        /// <param name="months">A number of months. The months parameter can be negative or positive.</param>
        /// <returns></returns>
        public DateTime AddMonths(int months)
        {
            return new DateTime(this.dateTimeObject.AddMonths(months));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public DateTime AddSeconds(double value)
        {
            return new DateTime(this.dateTimeObject.AddSeconds(value));
        }

        /// <summary>
        /// Returns a new DateTime that adds the specified number of ticks to the value of this instance.
        /// </summary>
        /// <param name="value">A number of 100-nanosecond ticks. The value parameter can be positive or negative.</param>
        /// <returns>An object whose value is the sum of the date and time represented by this instance and the time represented by value.</returns>
        public DateTime AddTicks(long value)
        {
            return new DateTime(this.dateTimeObject.AddTicks(value));
        }

        /// <summary>
        /// Returns a new DateTime that adds the specified number of years to the value of this instance.
        /// </summary>
        /// <param name="value">A number of years. The value parameter can be negative or positive.</param>
        /// <returns>An object whose value is the sum of the date and time represented by this instance and the number of years represented by value.</returns>
        public DateTime AddYears(int value)
        {
            return new DateTime(this.dateTimeObject.AddYears(value));
        }

        /// <summary>
        /// Compares two instances of DateTime and returns an integer that indicates whether the first instance is earlier than, the same as, or later than the second instance.
        /// </summary>
        /// <param name="t1">The first object to compare.</param>
        /// <param name="t2">The second object to compare.</param>
        /// <returns>A signed number indicating the relative values of t1 and t2.</returns>
        public int Compare(DateTime t1, DateTime t2)
        {
            return GSystem.DateTime.Compare(t1.dateTimeObject, t2.dateTimeObject);
        }

        /// <summary>
        /// Compares the value of this instance to a specified DateTime value and returns an integer that indicates whether this instance is earlier than, the same as, or later than the specified DateTime value.
        /// </summary>
        /// <param name="value">The object to compare to the current instance.</param>
        /// <returns>A signed number indicating the relative values of this instance and the value parameter.</returns>
        public int CompareTo(DateTime value)
        {
            return this.dateTimeObject.CompareTo(value.DateTimeObject);
        }
        /// <summary>
        /// Compares the value of this instance to a specified object that contains a specified DateTime value, and returns an integer that indicates whether this instance is earlier than, the same as, or later than the specified DateTime value.
        /// </summary>
        /// <param name="value">A boxed object to compare, or null.</param>
        /// <returns>A signed number indicating the relative values of this instance and value.</returns>
        /// 
        //TODO : Check implementation public int CompareTo2(object value)
        public int CompareTo2(object value)
        {
            const string Arg_MustBeDateTime = "Object must be of type DateTime.";

            if (value == null) return 1;
            if (!(value is DateTime))
            {
                throw new ArgumentException(Arg_MustBeDateTime);
            }
            return Compare(this, (DateTime)value);
        }

        /// <summary>
        /// Returns the number of days in the specified month and year.
        /// </summary>
        /// <param name="year">The year.</param>
        /// <param name="month">The month (a number ranging from 1 to 12).</param>
        /// <returns>The number of days in month for the specified year.</returns>
        public int DaysInMonth(int year, int month)
        {
            return GSystem.DateTime.DaysInMonth(year, month);
        }

        /// <summary>
        /// Returns a value indicating whether the value of this instance is equal to the value of the specified DateTime instance.
        /// </summary>
        /// <param name="value">The object to compare to this instance.</param>
        /// <returns>true if the value parameter equals the value of this instance; otherwise, false.</returns>
        public bool Equals(DateTime value)
        {
            return this.dateTimeObject.Equals(value.DateTimeObject);
        }

        /// <summary>
        /// Returns a value indicating whether this instance is equal to a specified object.
        /// </summary>
        /// <param name="value">The object to compare to this instance.</param>
        /// <returns>true if value is an instance of DateTime and equals the value of this instance; otherwise, false.</returns>
        // TODO : Check implementation
        public bool Equals2(object value)
        {
            return value is DateTime dt && this.dateTimeObject == dt.dateTimeObject;
        }

        /// <summary>
        /// Returns a value indicating whether two DateTime instances have the same date and time value.
        /// </summary>
        /// <param name="t1">The first object to compare.</param>
        /// <param name="t2">The second object to compare.</param>
        /// <returns>true if the two values are equal; otherwise, false.</returns>
        public bool Equals3(DateTime t1, DateTime t2)
        {
            return GSystem.DateTime.Equals(t1.dateTimeObject, t2.dateTimeObject);
        }

        /// <summary>
        /// Deserializes a 64-bit binary value and recreates an original serialized DateTime object.
        /// </summary>
        /// <param name="dateData">A 64-bit signed integer that encodes the Kind property in a 2-bit field and the Ticks property in a 62-bit field.</param>
        /// <returns>An object that is equivalent to the DateTime object that was serialized by the ToBinary() method.</returns>
        /// <exception cref="ArgumentException"> 
        /// <paramref name="dateData"/> is less than DateTime.MinValue or greater than DateTime.MaxValue.
        /// </exception>
        public DateTime FromBinary(long dateData)
        {
            return new DateTime(GSystem.DateTime.FromBinary(dateData));
        }

        /// <summary>
        /// Converts the specified Windows file time to an equivalent local time.
        /// </summary>
        /// <param name="fileTime">A Windows file time expressed in ticks.</param>
        /// <returns>An object that represents the local time equivalent of the date and time represented by the fileTime parameter.</returns>
        /// <exception cref="ArgumentOutOfRangeException"> 
        /// <paramref name="fileTime"/> is less than 0 or represents a time greater than DateTime.MaxValue.
        /// </exception> 
        public DateTime FromFileTime(long fileTime)
        {
            return new DateTime(GSystem.DateTime.FromFileTime(fileTime));
        }

        /// <summary>
        /// Converts the specified Windows file time to an equivalent UTC time.
        /// </summary>
        /// <param name="fileTime">A Windows file time expressed in ticks.</param>
        /// <returns>An object that represents the UTC time equivalent of the date and time represented by the fileTime parameter.</returns>
        /// <exception cref="ArgumentOutOfRangeException"> 
        /// <paramref name="fileTime"/> is less than 0 or represents a time greater than DateTime.MaxValue.
        /// </exception>         
        public DateTime FromFileTimeUtc(long fileTime)
        {
            return new DateTime(GSystem.DateTime.FromFileTimeUtc(fileTime));
        }

        /// <summary>
        /// Returns a DateTime equivalent to the specified OLE Automation Date.
        /// </summary>
        /// <param name="d">An OLE Automation Date value.</param>
        /// <returns>An object that represents the same date and time as d.</returns>
        /// <exception cref="ArgumentException"> 
        /// The date is not a valid OLE Automation Date value.
        /// </exception>        
        public DateTime FromOADate(double d)
        {
            return new DateTime(GSystem.DateTime.FromOADate(d));
        }

        /// <summary>
        /// Converts the value of this instance to all the string representations supported by the standard date and time format specifiers.
        /// </summary>
        /// <returns> A string array where each element is the representation of the value of this instance formatted with one of the standard date and time format specifiers.</returns>
        public string[] GetDateTimeFormats()
        {
            return this.dateTimeObject.GetDateTimeFormats();
        }

        /// <summary>
        /// Returns the hash code for this instance.
        /// </summary>
        /// <returns> A 32-bit signed integer hash code.</returns>
        public override int GetHashCode()
        {
            return this.dateTimeObject.GetHashCode();
        }

        /// <summary>
        /// Indicates whether this instance of DateTime is within the daylight saving time range for the current time zone.
        /// </summary>
        /// <returns> true if the value of the Kind property is Local or Unspecified and the value of this instance of DateTime is within the daylight saving time range for the local time zone; false if Kind is Utc.</returns>
        public bool IsDaylightSavingTime()
        {
            return this.dateTimeObject.IsDaylightSavingTime();
        }

        /// <summary>
        /// Returns an indication whether the specified year is a leap year
        /// </summary>
        /// <param name="year">A 4-digit year.</param>
        /// <returns> true if year is a leap year; otherwise, false.</returns>
        /// <exception cref="ArgumentOutOfRangeException"> 
        /// <paramref name="year"/> is less than 1 or greater than 9999. 
        /// </exception>
        public bool IsLeapYear(int year)
        {
            return GSystem.DateTime.IsLeapYear(year);
        }

        /// <summary>
        /// Converts the string representation of a date and time to its DateTime equivalent by using the conventions of the current culture.
        /// </summary>
        /// <param name="s">A string that contains a date and time to convert. See The string to parse for more information.</param>
        /// <returns>An object that is equivalent to the date and time contained in s.</returns>
        /// <exception cref="ArgumentNullException"> 
        /// <paramref name="s"/> is null.
        /// </exception>
        public DateTime Parse(string s)
        {
            return new DateTime(GSystem.DateTime.Parse(s));
        }

        public DateTime Parse2(string s, GSystem.IFormatProvider provider)
        {
            return new DateTime(GSystem.DateTime.Parse(s, provider));
        }

        public DateTime Parse3(string s, IFormatProvider provider, GSystem.Globalization.DateTimeStyles styles)
        {
            return new DateTime(GSystem.DateTime.Parse(s, provider, styles));
        }

        public DateTime ParseExact(string s, string format, IFormatProvider provider)
        {
            return new DateTime(GSystem.DateTime.ParseExact(s, format, provider));

        }

        public DateTime ParseExact2(string s, string format, IFormatProvider provider, GSystem.Globalization.DateTimeStyles style)
        {
            return new DateTime(GSystem.DateTime.ParseExact(s, format, provider, style));
        }

        public DateTime ParseExact3(string s, [In] ref string[] formats, IFormatProvider provider, GSystem.Globalization.DateTimeStyles style)
        {
            return new DateTime(GSystem.DateTime.ParseExact(s, formats, provider, style));
        }

        /// <summary>
        /// Creates a new DateTime object that has the same number of ticks as the specified DateTime, but is designated as either local time, Coordinated Universal Time (UTC), or neither, as indicated by the specified DateTimeKind value.
        /// </summary>
        /// <param name="value">A date and time.</param>
        /// <param name="kind">One of the enumeration values that indicates whether the new object represents local time, UTC, or neither.</param>
        /// <returns>A new object that has the same number of ticks as the object represented by the value parameter and the DateTimeKind value specified by the kind parameter.</returns>
        public DateTime SpecifyKind(DateTime value, DateTimeKind kind)
        {
            return new DateTime(GSystem.DateTime.SpecifyKind(value.dateTimeObject, (GSystem.DateTimeKind)kind));
        }

        public DateTime Subtract(TimeSpan value)
        {
            return new DateTime(this.dateTimeObject.Subtract(value.TimeSpanObject));
        }

        public TimeSpan Subtract2(DateTime value)
        {
            return new TimeSpan(this.dateTimeObject.Subtract(value.dateTimeObject));
        }

        public long ToBinary()
        {
            return this.dateTimeObject.ToBinary();
        }

        public long ToFileTime()
        {
            return this.dateTimeObject.ToFileTime();
        }

        public long ToFileTimeUtc()
        {
            return this.dateTimeObject.ToFileTimeUtc();
        }

        /// <summary>
        /// Converts the value of the current DateTime object to local time.
        /// </summary>
        /// <returns>An object whose Kind property is Local, and whose value is the local time equivalent to the value of the current DateTime object, or DateTime.MaxValue if the converted value is too large to be represented by a DateTime object, or DateTime.MinValue if the converted value is too small to be represented as a DateTime object.</returns>        
        public DateTime ToLocalTime()
        {
            return new DateTime(this.dateTimeObject.ToLocalTime());
        }

        public string ToLongDateString()
        {
            return this.dateTimeObject.ToLongDateString();
        }

        public string ToLongTimeString()
        {
            return this.dateTimeObject.ToLongTimeString();
        }

        /// <summary>
        /// Converts the value of this instance to the equivalent OLE Automation date.
        /// </summary>
        /// <returns>A double-precision floating-point number that contains an OLE Automation date equivalent to the value of this instance.</returns>
        /// <exception cref="OverflowException"> The value of this instance cannot be represented as an OLE Automation Date
        /// </exception>
        public double ToOADate()
        {
            return this.dateTimeObject.ToOADate();
        }

        public string ToShortDateString()
        {
            return this.dateTimeObject.ToShortDateString();
        }

        public string ToShortTimeString()
        {
            return this.dateTimeObject.ToShortTimeString();
        }

        public override string ToString()
        {
            return this.dateTimeObject.ToString();
        }

        /// <summary>
        /// Converts the value of the current DateTime object to its equivalent string representation using the specified format and the formatting conventions of the current culture.
        /// </summary>
        /// <param name="format">A standard or custom date and time format string.</param>
        /// <returns>A string representation of value of the current DateTime object as specified by format.</returns>
        /// <exception cref="FormatException">
        /// <paramref name="format"/> The length of format is 1, and it is not one of the format specifier characters defined for DateTimeFormatInfo.
        ///
        /// -or-
        ///
        /// <paramref name="format"/> does not contain a valid custom format pattern
        /// </exception>
        /// <exception cref="ArgumentOutOfRangeException"> 
        /// The date and time is outside the range of dates supported by the calendar used by the current culture.
        /// </exception>
        public string ToString2(string format)
        {
            return this.dateTimeObject.ToString(format);
        }

        public string ToString3(IFormatProvider provider)
        {
            return this.dateTimeObject.ToString(provider);
        }

        public string ToString4(string format, IFormatProvider provider)
        {
            return this.dateTimeObject.ToString(format, provider);
        }

        /// <summary>
        /// Converts the value of the current DateTime object to Coordinated Universal Time (UTC).
        /// </summary>
        /// <returns> An object whose Kind property is Utc, and whose value is the UTC equivalent to the value of the current DateTime object, or DateTime.MaxValue if the converted value is too large to be represented by a DateTime object, or DateTime.MinValue if the converted value is too small to be represented by a DateTime object.</returns>
        public DateTime ToUniversalTime()
        {
            return new DateTime(this.dateTimeObject.ToUniversalTime());
        }

        public bool TryParse(string s, out DateTime result)
        {
            bool tryParse = GSystem.DateTime.TryParse(s, out GSystem.DateTime outResult);
            result = new DateTime(outResult);
            return tryParse;
        }

        public bool TryParse2(string s, IFormatProvider provider, GSystem.Globalization.DateTimeStyles styles, out DateTime result)
        {
            bool tryParse = GSystem.DateTime.TryParse(s, provider, styles, out GSystem.DateTime outResult);
            result = new DateTime(outResult);
            return tryParse;
        }

        //Operators

        public DateTime Addition(DateTime d, TimeSpan t)
        {
            return new DateTime(d.DateTimeObject + t.TimeSpanObject);
        }

        public bool Equality(DateTime d1, DateTime d2)
        { 
            return (d1.DateTimeObject == d2.DateTimeObject); 
        }

        public bool GreaterThan(DateTime t1, DateTime t2)
        {
            return (t1.DateTimeObject > t2.DateTimeObject);
        }

        public bool GreaterThanOrEqual(DateTime t1, DateTime t2)
        { 
            return (t1.DateTimeObject >= t2.DateTimeObject);
        }
        public bool Inequality(DateTime t1, DateTime t2)
        {
            return (t1.DateTimeObject != t2.DateTimeObject);
        }

        public bool LessThan(DateTime t1, DateTime t2)
        {
            return (t1.DateTimeObject < t2.DateTimeObject);
        }
        public bool LessThanOrEqual(DateTime t1, DateTime t2)
        {
            return (t1.DateTimeObject <= t2.DateTimeObject);
        }

        public TimeSpan Subtraction(DateTime d1, DateTime d2)
        {
            return new TimeSpan(d1.DateTimeObject - d2.DateTimeObject);
        }

        public DateTime Subtraction2(DateTime d, TimeSpan t)
        {
            return new DateTime(d.DateTimeObject - t.TimeSpanObject);
        }

    }
}

// https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1
// https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=net-7.0
// https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/DateTime.cs
// https://learn.microsoft.com/en-us/dotnet/core/native-interop/expose-components-to-com
// https://www.codeproject.com/Articles/3511/Exposing-NET-Components-to-COM
// https://stackoverflow.com/questions/2714430/why-should-i-not-use-autodual
// https://www.red-gate.com/simple-talk/development/dotnet-development/build-and-deploy-a-net-com-assembly/
// https://www.thevbahelp.com/post/calling-c-sharp-code-from-vba-com-interop#viewer-1hj2h

