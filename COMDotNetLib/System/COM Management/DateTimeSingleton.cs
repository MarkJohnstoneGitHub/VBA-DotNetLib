using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.SqlServer.Server;

namespace DotNetLib.System.COM_Management
{
    [ComVisible(true)]
    [Guid("9614E02F-B631-44C5-B2AA-31A87868A9A2")]
    [ProgId("DotNetLib.System.DateTimeSingleton")]
    [Description("DateTime factory methods and static members.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDateTimeSingleton))]
    public class DateTimeSingleton : IDateTimeSingleton
    {

        // Factory Methods

        /// <summary>
        /// Initializes a new instance of the DateTime structure to a specified number of pTicks and to Coordinated Universal Time (UTC) or local time.
        /// </summary>
        /// <param name="pTicks">A date and time expressed in the number of 100-nanosecond intervals that have elapsed since January 1, 0001 at 00:00:00.000 in the Gregorian calendar.</param>
        /// <param name="pKind">One of the enumeration values that indicates whether pTicks specifies a local time, Coordinated Universal Time (UTC), or neither.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException"> 
        /// pTicks is less than DateTime.MinValue or greater than DateTime.MaxValue.
        /// </exception>
        /// <exception cref="ArgumentException"> 
        /// pKind is not one of the DateTimeKind values.
        /// </exception>
        /// 
        public DateTime CreateFromTicks(long pTicks, DateTimeKind pKind = DateTimeKind.Unspecified)
        {
            return new DateTime(pTicks, pKind);
        }
        public DateTime CreateFromDate(int pYear, int pMonth, int pDay)
        {
            return new DateTime(pYear, pMonth, pDay);
        }

        /// <summary>
        /// Initializes a new instance of the DateTime structure to the specified pYear, pMonth, pDay, pHour, pMinute, pSecond, and pMillisecond.
        /// </summary>
        /// <param name="pYear">The pYear (1 through 9999).</param>
        /// <param name="pMonth">The pMonth (1 through 12).</param>
        /// <param name="pDay">The pDay (1 through the number of days in pMonth).</param>
        /// <param name="pHour">The hours (0 through 23).</param>
        /// <param name="pMinute">The minutes (0 through 59).</param>
        /// <param name="pSecond">The seconds (0 through 59).</param>
        /// <param name="pMillisecond">The milliseconds (0 through 999).</param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException">
        /// <paramref name="pYear"/> is less than 1 or greater than 9999.
        ///
        /// -or-
        ///
        /// <paramref name="pMonth"/> is less than 1 or greater than 12.
        ///
        /// -or-
        ///
        /// <paramref name="pDay"/> is less than 1 or greater than the number of days in <paramref name="pMonth"/>.
        ///
        /// -or-
        ///
        /// <paramref name="pHour"/> is less than 0 or greater than 23.
        ///
        /// -or-
        ///
        /// <paramref name="pMinute"/> is less than 0 or greater than 59.
        ///
        /// -or-
        ///
        /// <paramref name="pSecond"/> is less than 0 or greater than 59.
        ///
        /// -or-
        ///
        /// <paramref name="pMillisecond"/> is less than 0 or greater than 999.
        /// </exception>        
        public DateTime CreateFromDateTime(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, int pMillisecond = 0)
        {
            return new DateTime(pYear, pMonth, pDay, pHour, pMinute, pSecond, pMillisecond);
        }

        public DateTime CreateFromDateTimeKind(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, DateTimeKind pKind)
        {
            return new DateTime(pYear, pMonth, pDay, pHour, pMinute, pSecond, pKind);
        }

        public DateTime CreateFromDateTimeKind2(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, int pMillisecond, DateTimeKind pKind)
        {
            return new DateTime(pYear, pMonth, pDay, pHour, pMinute, pSecond, pMillisecond, pKind);
        }

        //Fields

        /// <summary>
        /// Represents the largest possible value of DateTime. This field is read-only.
        /// </summary>
        public DateTime MaxValue => DateTime.MaxValue;

        /// <summary>
        /// Represents the smallest possible value of DateTime. This field is read-only.
        /// </summary>
        public DateTime MinValue => DateTime.MinValue; 

        // Properties

        /// <summary>
        /// The value of this constant is equivalent to 00:00:00.0000000 UTC, January 1, 1970, in the Gregorian calendar. UnixEpoch defines the point in time when Unix time is equal to 0.
        /// </summary>
        //public DateTime UnixEpoch => dtUnixEpoch;

        /// <summary>
        /// The nanoseconds component, expressed as a value between 0 and 900 (in increments of 100 nanoseconds).
        /// </summary>
        //public int Nanosecond => this.objDateTime.Nanosecond;

        /// <summary>
        /// Gets a DateTime object that is set to the current date and time on this computer, expressed as the local time.
        /// </summary>
        public DateTime Now => DateTime.Now;

        /// <summary>
        /// Gets the current date.
        /// </summary>
        public DateTime Today => DateTime.Today;

        /// <summary>
        /// Gets a DateTime object that is set to the current date and time on this computer, expressed as the Coordinated Universal Time (UTC).
        /// </summary>
        public DateTime UtcNow => DateTime.UtcNow;


        // Methods

        /// <summary>
        /// Compares two instances of DateTime and returns an integer that indicates whether the first instance is earlier than, the same as, or later than the pSecond instance.
        /// </summary>
        /// <param name="t1">The first object to compare.</param>
        /// <param name="t2">The pSecond object to compare.</param>
        /// <returns>A signed number indicating the relative values of t1 and t2.</returns>
        public int Compare(DateTime t1, DateTime t2)
        {
            return DateTime.Compare(t1, t2);
        }

        /// <summary>
        /// Returns the number of days in the specified pMonth and pYear.
        /// </summary>
        /// <param name="year">The pYear.</param>
        /// <param name="month">The pMonth (a number ranging from 1 to 12).</param>
        /// <returns>The number of days in pMonth for the specified pYear.</returns>
        public int DaysInMonth(int year, int month)
        {
            return DateTime.DaysInMonth(year, month);
        }

        /// <summary>
        /// Returns a value indicating whether two DateTime instances have the same date and time value.
        /// </summary>
        /// <param name="t1">The first object to compare.</param>
        /// <param name="t2">The pSecond object to compare.</param>
        /// <returns>true if the two values are equal; otherwise, false.</returns>
        public bool Equals(DateTime t1, DateTime t2)
        {
            return DateTime.Equals(t1, t2); 
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
            return DateTime.FromBinary(dateData); 
        }

        /// <summary>
        /// Converts the specified Windows file time to an equivalent local time.
        /// </summary>
        /// <param name="fileTime">A Windows file time expressed in pTicks.</param>
        /// <returns>An object that represents the local time equivalent of the date and time represented by the fileTime parameter.</returns>
        /// <exception cref="ArgumentOutOfRangeException"> 
        /// <paramref name="fileTime"/> is less than 0 or represents a time greater than DateTime.MaxValue.
        /// </exception> 
        public DateTime FromFileTime(long fileTime)
        {
            return DateTime.FromFileTime(fileTime); 
        }

        /// <summary>
        /// Converts the specified Windows file time to an equivalent UTC time.
        /// </summary>
        /// <param name="fileTime">A Windows file time expressed in pTicks.</param>
        /// <returns>An object that represents the UTC time equivalent of the date and time represented by the fileTime parameter.</returns>
        /// <exception cref="ArgumentOutOfRangeException"> 
        /// <paramref name="fileTime"/> is less than 0 or represents a time greater than DateTime.MaxValue.
        /// </exception>         
        public DateTime FromFileTimeUtc(long fileTime)
        {
            return DateTime.FromFileTime(fileTime); 
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
            return DateTime.FromOADate(d);
        }

        /// <summary>
        /// Returns an indication whether the specified pYear is a leap pYear
        /// </summary>
        /// <param name="year">A 4-digit pYear.</param>
        /// <returns> true if pYear is a leap pYear; otherwise, false.</returns>
        /// <exception cref="ArgumentOutOfRangeException"> 
        /// <paramref name="year"/> is less than 1 or greater than 9999. 
        /// </exception>
        public bool IsLeapYear(int year)
        {
            return DateTime.IsLeapYear(year);
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
            return DateTime.Parse(s);
        }

        public DateTime Parse2(string s, GSystem.IFormatProvider provider)
        {
            return DateTime.Parse(s, provider); 
        }

        public DateTime Parse3(string s, IFormatProvider provider, GSystem.Globalization.DateTimeStyles styles)
        {
            return DateTime.Parse(s, provider, styles);
        }

        public DateTime ParseExact(string s, string format, IFormatProvider provider)
        {
            return DateTime.ParseExact(s, format, provider); 

        }

        public DateTime ParseExact2(string s, string format, IFormatProvider provider, GSystem.Globalization.DateTimeStyles style)
        {
            return DateTime.ParseExact(s, format, provider, style);  
        }

        public DateTime ParseExact3(string s, [In] ref string[] formats, IFormatProvider provider, GSystem.Globalization.DateTimeStyles style)
        {
            return DateTime.ParseExact(s, ref formats, provider, style);
        }

        /// <summary>
        /// Creates a new DateTime object that has the same number of pTicks as the specified DateTime, but is designated as either local time, Coordinated Universal Time (UTC), or neither, as indicated by the specified DateTimeKind value.
        /// </summary>
        /// <param name="value">A date and time.</param>
        /// <param name="pKind">One of the enumeration values that indicates whether the new object represents local time, UTC, or neither.</param>
        /// <returns>A new object that has the same number of pTicks as the object represented by the value parameter and the DateTimeKind value specified by the pKind parameter.</returns>
        public DateTime SpecifyKind(DateTime value, DateTimeKind pKind)
        {
            return DateTime.SpecifyKind(value, pKind);
        }

        public bool TryParse(string s, out DateTime result)
        {
            return DateTime.TryParse(s, out result);
        }

        public bool TryParse2(string s, IFormatProvider provider, GSystem.Globalization.DateTimeStyles styles, out DateTime result)
        {
            return DateTime.TryParse(s, provider, styles, out result);
        }

        public bool TryParseExact(string s, string format, IFormatProvider provider, GSystem.Globalization.DateTimeStyles style, out DateTime result)
        {
            return DateTime.TryParseExact(s, format, provider, style, out result);
        }

        public bool TryParseExact2(string s, [In] ref string[] formats, IFormatProvider provider, GSystem.Globalization.DateTimeStyles style, out DateTime result)
        {
            return DateTime.TryParseExact(s, formats, provider, style, out result);
        }

        //Operators

        public DateTime Addition(DateTime dt, TimeSpan ts)
        {
            return DateTime.Addition(dt, ts); 
        }

        public bool Equality(DateTime d1, DateTime d2)
        {
            return DateTime.Equality(d1, d2);
        }

        public bool GreaterThan(DateTime t1, DateTime t2)
        {
            return DateTime.GreaterThan(t1, t2);
        }

        public bool GreaterThanOrEqual(DateTime t1, DateTime t2)
        {
            return DateTime.GreaterThanOrEqual(t1, t2);
        }
        public bool Inequality(DateTime t1, DateTime t2)
        {
            return DateTime.Inequality(t1, t2);
        }

        public bool LessThan(DateTime t1, DateTime t2)
        {
            return DateTime.LessThan(t1, t2);
        }
        public bool LessThanOrEqual(DateTime t1, DateTime t2)
        {
            return DateTime.LessThanOrEqual(t1, t2);
        }

        public TimeSpan Subtraction(DateTime d1, DateTime d2)
        {
            return DateTime.Subtraction(d1, d2);
        }

        public DateTime Subtraction2(DateTime d, TimeSpan t)
        {
            return DateTime.Subtraction(d, t);
        }

    }
}
