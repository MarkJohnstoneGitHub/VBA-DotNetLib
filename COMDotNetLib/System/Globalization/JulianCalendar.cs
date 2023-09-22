// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.juliancalendar?view=netframework-4.8.1

using GGlobalization = global::System.Globalization;
using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.Extensions;
using System.Globalization;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("970BBA5C-E584-4C8E-904C-D540C19A940B")]
    [ProgId("DotNetLib.System.Globalization.JulianCalendar")]
    [Description("Represents the Julian calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IJulianCalendar))]
    public class JulianCalendar : IJulianCalendar, Calendar, IWrappedObject
    {
        private GGlobalization.JulianCalendar _julianCalendar;

        //Constructors
        public JulianCalendar()
        {
            _julianCalendar = new GGlobalization.JulianCalendar();
        }

        public JulianCalendar(GGlobalization.JulianCalendar julianCalendar)
        {
            _julianCalendar = julianCalendar;
        }

        // Properties
        public GGlobalization.JulianCalendar WrappedJulianCalendar => _julianCalendar;

        public object WrappedObject => _julianCalendar;

        public CalendarAlgorithmType AlgorithmType => _julianCalendar.AlgorithmType;

        public int[] Eras => _julianCalendar.Eras;

        public bool IsReadOnly => _julianCalendar.IsReadOnly;

        public DateTime MaxSupportedDateTime => new DateTime(_julianCalendar.MaxSupportedDateTime);

        public DateTime MinSupportedDateTime => new DateTime(_julianCalendar.MinSupportedDateTime);

        public int TwoDigitYearMax
        {
            get => _julianCalendar.TwoDigitYearMax;
            set => _julianCalendar.TwoDigitYearMax = value;
        }

        // Methods

        public DateTime AddDays(DateTime time, int days)
        {
            return new DateTime(_julianCalendar.AddDays(time.WrappedDateTime, days));
        }

        public DateTime AddHours(DateTime time, int hours)
        {
            return new DateTime(_julianCalendar.AddHours(time.WrappedDateTime, hours));
        }

        public DateTime AddMilliseconds(DateTime time, double milliseconds)
        {
            return new DateTime(_julianCalendar.AddMilliseconds(time.WrappedDateTime, milliseconds));
        }

        public DateTime AddMinutes(DateTime time, int minutes)
        {
            return new DateTime(_julianCalendar.AddMilliseconds(time.WrappedDateTime, minutes));
        }

        public DateTime AddMonths(DateTime time, int months)
        {
            return new DateTime(_julianCalendar.AddMilliseconds(time.WrappedDateTime, months));
        }

        public DateTime AddSeconds(DateTime time, int seconds)
        {
            return new DateTime(_julianCalendar.AddSeconds(time.WrappedDateTime, seconds));
        }

        public DateTime AddWeeks(DateTime time, int weeks)
        {
            return new DateTime(_julianCalendar.AddWeeks(time.WrappedDateTime, weeks));
        }

        public DateTime AddYears(DateTime time, int years)
        {
            return new DateTime(_julianCalendar.AddYears(time.WrappedDateTime, years));
        }

        public object Clone()
        {
            return new JapaneseCalendar((GGlobalization.JapaneseCalendar)_julianCalendar.Clone());
        }

        public override bool Equals(object obj)
        {
            return obj is JapaneseCalendar cal && _julianCalendar.Equals(cal.WrappedJapaneseCalendar);
        }

        public int GetDayOfMonth(DateTime time)
        {
            return _julianCalendar.GetDayOfMonth(time.WrappedDateTime);
        }

        public GSystem.DayOfWeek GetDayOfWeek(DateTime time)
        {
            return _julianCalendar.GetDayOfWeek(time.WrappedDateTime);
        }

        public int GetDayOfYear(DateTime time)
        {
            return _julianCalendar.GetDayOfYear(time.WrappedDateTime);
        }

        public int GetDaysInMonth(int year, int month, int era = 0)
        {
            return _julianCalendar.GetDaysInMonth(year, month, era);
        }

        public int GetDaysInYear(int year, int era = 0)
        {
            return _julianCalendar.GetDaysInYear(year, era);
        }

        public int GetEra(DateTime time)
        {
            return _julianCalendar.GetEra(time.WrappedDateTime);
        }

        public override int GetHashCode()
        {
            return _julianCalendar.GetHashCode();
        }

        public int GetHour(DateTime time)
        {
            return _julianCalendar.GetHour(time.WrappedDateTime);
        }

        public int GetLeapMonth(int year, int era = 0)
        {
            return _julianCalendar.GetLeapMonth(year, era);
        }

        public double GetMilliseconds(DateTime time)
        {
            return _julianCalendar.GetMilliseconds(time.WrappedDateTime);
        }

        public int GetMinute(DateTime time)
        {
            return _julianCalendar.GetMinute(time.WrappedDateTime);
        }

        public int GetMonth(DateTime time)
        {
            return _julianCalendar.GetMonth(time.WrappedDateTime);
        }

        public int GetMonthsInYear(int year, int era = 0)
        {
            return _julianCalendar.GetMonthsInYear(year, era);
        }

        public int GetSecond(DateTime time)
        {
            return _julianCalendar.GetSecond(time.WrappedDateTime);
        }

        public int GetWeekOfYear(DateTime time, CalendarWeekRule rule, GSystem.DayOfWeek firstDayOfWeek)
        {
            return _julianCalendar.GetWeekOfYear(time.WrappedDateTime, rule, firstDayOfWeek);
        }

        public int GetYear(DateTime time)
        {
            return _julianCalendar.GetYear(time.WrappedDateTime);
        }

        public bool IsLeapDay(int year, int month, int day, int era = 0)
        {
            return _julianCalendar.IsLeapDay(year, month, day, era);
        }

        public bool IsLeapMonth(int year, int month, int era = 0)
        {
            return _julianCalendar.IsLeapMonth(year, month, era);
        }

        public bool IsLeapYear(int year, int era = 0)
        {
            return _julianCalendar.IsLeapYear(year, era);
        }

        public DateTime ToDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond, int era = 0)
        {
            return new DateTime(_julianCalendar.ToDateTime(year, month, day, hour, minute, second, millisecond, era));
        }

        public int ToFourDigitYear(int year)
        {
            return _julianCalendar.ToFourDigitYear(year);
        }

        new public string ToString()
        { 
            return _julianCalendar.ToString(); 
        }
    }
}
