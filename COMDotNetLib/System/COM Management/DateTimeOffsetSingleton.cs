// https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset?view=netframework-4.8.1

using GSystem = global::System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;

namespace DotNetLib.System
{

    [ComVisible(true)]
    [Description("DateTimeOffset factory methods and static members.")]
    [Guid("E7D6FD84-B6B3-450D-8D19-EEAED008F579")]
    [ProgId("DotNetLib.System.DateTimeOffsetSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDateTimeOffsetSingleton))]
    public class DateTimeOffsetSingleton : IDateTimeOffsetSingleton
    {
        // Factory Methods
        public DateTimeOffset CreateFromDateTime(DateTime pDateTime)
        {
            return new DateTimeOffset(pDateTime);
        }

        public DateTimeOffset CreateFromDateTime2(DateTime pDateTime, TimeSpan pOffset)
        {
            return new DateTimeOffset(pDateTime, pOffset);
        }

        public DateTimeOffset CreateFromDateTimeParts(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, TimeSpan pOffset)
        {
            return new DateTimeOffset(pYear, pMonth, pDay, pHour, pMinute, pSecond, pOffset);
        }

        public DateTimeOffset CreateFromDateTimeParts2(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, int pMillisecond, TimeSpan pOffset)
        {
            return new DateTimeOffset(pYear, pMonth, pDay, pHour, pMinute, pSecond, pMillisecond, pOffset);
        }

        public DateTimeOffset CreateFromDateTimeParts3(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, int pMillisecond, GSystem.Globalization.Calendar pCalendar, TimeSpan pOffset)
        {
            return new DateTimeOffset(pYear, pMonth, pDay, pHour, pMinute, pSecond, pMillisecond, pCalendar, pOffset);
        }

        public DateTimeOffset CreateFromTicks(long pTicks, TimeSpan pOffset)
        {
            return new DateTimeOffset(pTicks, pOffset);
        }

        // Fields

        public DateTimeOffset MaxValue => DateTimeOffset.MaxValue;
        public DateTimeOffset MinValue => DateTimeOffset.MinValue;

        // Properties

        public DateTimeOffset Now => DateTimeOffset.Now;

        public DateTimeOffset UtcNow => DateTimeOffset.UtcNow;

        // Methods

        public int Compare(DateTimeOffset first, DateTimeOffset second)
        {
            return DateTimeOffset.Compare(first, second);
        }

        public bool Equals(DateTimeOffset first, DateTimeOffset second)
        {
            return DateTimeOffset.Equals(first, second);
        }

        public DateTimeOffset FromFileTime(long fileTime)
        {
            return DateTimeOffset.FromFileTime(fileTime);
        }

        public DateTimeOffset FromUnixTimeMilliseconds(long pMilliseconds)
        {
            return DateTimeOffset.FromUnixTimeMilliseconds(pMilliseconds);
        }

        public DateTimeOffset FromUnixTimeSeconds(long pSeconds)
        {
            return DateTimeOffset.FromUnixTimeSeconds(pSeconds);
        }

        public DateTimeOffset Parse(string input)
        {
            return DateTimeOffset.Parse(input);
        }

        public DateTimeOffset Parse2(string input, IFormatProvider formatProvider)
        {
            return new DateTimeOffset(GSystem.DateTimeOffset.Parse(input, formatProvider));
        }

        public DateTimeOffset Parse3(string input, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles)
        {
            return DateTimeOffset.Parse(input, formatProvider, styles);
        }

        public DateTimeOffset ParseExact(string input, string format, IFormatProvider formatProvider)
        {
            return DateTimeOffset.ParseExact(input, format, formatProvider);
        }

        public DateTimeOffset ParseExact2(string input, string format, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles)
        {
            return DateTimeOffset.ParseExact(input, format, formatProvider, styles);
        }

        public DateTimeOffset ParseExact3(string input, [In] ref string[] formats, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles)
        {
            return DateTimeOffset.ParseExact(input,ref formats, formatProvider, styles);
        }

        public bool TryParse(string input, out DateTimeOffset result)
        {
            return DateTimeOffset.TryParse(input, out result);
        }

        public bool TryParse2(string input, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles, out DateTimeOffset result)
        {
            return DateTimeOffset.TryParse(input, formatProvider, styles, out result);
        }

        public bool TryParseExact(string input, string format, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles, out DateTimeOffset result)
        {
            return DateTimeOffset.TryParseExact(input, format, formatProvider, styles, out result);

        }

        public bool TryParseExact2(string input, [In] ref string[] formats, IFormatProvider formatProvider, GSystem.Globalization.DateTimeStyles styles, out DateTimeOffset result)
        {
            return DateTimeOffset.TryParseExact(input, formats, formatProvider, styles, out result);
        }

        // Operators

        public DateTimeOffset Addition(DateTimeOffset dateTimeOffset, TimeSpan timeSpan)
        {
            return  DateTimeOffset.Addition(dateTimeOffset, timeSpan);
        }

        public bool Equality(DateTimeOffset left, DateTimeOffset right)
        {
            return DateTimeOffset.Equality(left, right);
        }

        public bool GreaterThan(DateTimeOffset left, DateTimeOffset right)
        {
            return DateTimeOffset.GreaterThan(left, right);
        }

        public bool GreaterThanOrEqual(DateTimeOffset left, DateTimeOffset right)
        {
            return DateTimeOffset.GreaterThanOrEqual(left, right);
        }

        public DateTimeOffset Implicit(DateTime dateTime)
        {
            return DateTimeOffset.Implicit(dateTime);
        }

        public bool Inequality(DateTimeOffset left, DateTimeOffset right)
        {
            return DateTimeOffset.Inequality(left, right);
        }

        public bool LessThan(DateTimeOffset left, DateTimeOffset right)
        {
            return DateTimeOffset.LessThan(left,right);
        }

        public bool LessThanOrEqual(DateTimeOffset left, DateTimeOffset right)
        {
            return DateTimeOffset.LessThanOrEqual(left, right);
        }

        public TimeSpan Subtraction(DateTimeOffset left, DateTimeOffset right)
        {
            return DateTimeOffset.Subtraction(left,right);
        }

        public DateTimeOffset Subtraction2(DateTimeOffset dateTimeOffset, TimeSpan timeSpan)
        {
            return DateTimeOffset.Subtraction(dateTimeOffset, timeSpan);
        }
    }
}
