// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.hebrewcalendar?view=netframework-4.8.1

using GSystem = global::System;
using GGlobalization = global::System.Globalization;
using System;
using System.Globalization;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.Extensions;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("3857D80C-06E8-466B-BD62-4602037D7299")]
    [ProgId("DotNetLib.System.Globalization.HebrewCalendar")]
    [Description("Represents the Hebrew calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IHebrewCalendar))]
    public class HebrewCalendar : IHebrewCalendar, Calendar, IWrappedObject
    {
        private GGlobalization.HebrewCalendar _hebrewCalendar;

        //Constructors
        public HebrewCalendar()
        {
            _hebrewCalendar = new GGlobalization.HebrewCalendar();
        }

        public HebrewCalendar(GGlobalization.HebrewCalendar hebrewCalendar)
        {
            _hebrewCalendar = hebrewCalendar;
        }

        // Properties
        public GGlobalization.HebrewCalendar WrappedHebrewCalendar => _hebrewCalendar;

        public object WrappedObject => _hebrewCalendar;

        public CalendarAlgorithmType AlgorithmType => _hebrewCalendar.AlgorithmType;

        public int[] Eras => _hebrewCalendar.Eras;

        public bool IsReadOnly => _hebrewCalendar.IsReadOnly;

        public DateTime MaxSupportedDateTime => new DateTime(_hebrewCalendar.MaxSupportedDateTime);

        public DateTime MinSupportedDateTime => new DateTime(_hebrewCalendar.MinSupportedDateTime);

        public int TwoDigitYearMax
        {
            get => _hebrewCalendar.TwoDigitYearMax;
            set => _hebrewCalendar.TwoDigitYearMax = value;
        }

        // Methods

        public DateTime AddDays(DateTime time, int days)
        {
            return new DateTime(_hebrewCalendar.AddDays(time.WrappedDateTime, days));
        }

        public DateTime AddHours(DateTime time, int hours)
        {
            return new DateTime(_hebrewCalendar.AddHours(time.WrappedDateTime, hours));
        }

        public DateTime AddMilliseconds(DateTime time, double milliseconds)
        {
            return new DateTime(_hebrewCalendar.AddMilliseconds(time.WrappedDateTime, milliseconds));
        }

        public DateTime AddMinutes(DateTime time, int minutes)
        {
            return new DateTime(_hebrewCalendar.AddMilliseconds(time.WrappedDateTime, minutes));
        }

        public DateTime AddMonths(DateTime time, int months)
        {
            return new DateTime(_hebrewCalendar.AddMilliseconds(time.WrappedDateTime, months));
        }

        public DateTime AddSeconds(DateTime time, int seconds)
        {
            return new DateTime(_hebrewCalendar.AddSeconds(time.WrappedDateTime, seconds));
        }

        public DateTime AddWeeks(DateTime time, int weeks)
        {
            return new DateTime(_hebrewCalendar.AddWeeks(time.WrappedDateTime, weeks));
        }

        public DateTime AddYears(DateTime time, int years)
        {
            return new DateTime(_hebrewCalendar.AddYears(time.WrappedDateTime, years));
        }

        public object Clone()
        {
            return new HebrewCalendar((GGlobalization.HebrewCalendar)_hebrewCalendar.Clone());
        }

        public override bool Equals(object obj)
        {
            return obj is HebrewCalendar cal && _hebrewCalendar.Equals(cal.WrappedHebrewCalendar);
        }

        public int GetDayOfMonth(DateTime time)
        {
            return _hebrewCalendar.GetDayOfMonth(time.WrappedDateTime);
        }

        public GSystem.DayOfWeek GetDayOfWeek(DateTime time)
        {
            return _hebrewCalendar.GetDayOfWeek(time.WrappedDateTime);
        }

        public int GetDayOfYear(DateTime time)
        {
            return _hebrewCalendar.GetDayOfYear(time.WrappedDateTime);
        }

        public int GetDaysInMonth(int year, int month, int era = 0)
        {
            return _hebrewCalendar.GetDaysInMonth(year, month, era);
        }

        public int GetDaysInYear(int year, int era = 0)
        {
            return _hebrewCalendar.GetDaysInYear(year, era);
        }

        public int GetEra(DateTime time)
        {
            return _hebrewCalendar.GetEra(time.WrappedDateTime);
        }

        public override int GetHashCode()
        {
            return _hebrewCalendar.GetHashCode();
        }

        public int GetHour(DateTime time)
        {
            return _hebrewCalendar.GetHour(time.WrappedDateTime);
        }

        public int GetLeapMonth(int year, int era = 0)
        {
            return _hebrewCalendar.GetLeapMonth(year, era);
        }

        public double GetMilliseconds(DateTime time)
        {
            return _hebrewCalendar.GetMilliseconds(time.WrappedDateTime);
        }

        public int GetMinute(DateTime time)
        {
            return _hebrewCalendar.GetMinute(time.WrappedDateTime);
        }

        public int GetMonth(DateTime time)
        {
            return _hebrewCalendar.GetMonth(time.WrappedDateTime);
        }

        public int GetMonthsInYear(int year, int era = 0)
        {
            return _hebrewCalendar.GetMonthsInYear(year, era);
        }

        public int GetSecond(DateTime time)
        {
            return _hebrewCalendar.GetSecond(time.WrappedDateTime);
        }

        public int GetWeekOfYear(DateTime time, CalendarWeekRule rule, GSystem.DayOfWeek firstDayOfWeek)
        {
            return _hebrewCalendar.GetWeekOfYear(time.WrappedDateTime, rule, firstDayOfWeek);
        }

        public int GetYear(DateTime time)
        {
            return _hebrewCalendar.GetYear(time.WrappedDateTime);
        }

        public bool IsLeapDay(int year, int month, int day, int era = 0)
        {
            return _hebrewCalendar.IsLeapDay(year, month, day, era);
        }

        public bool IsLeapMonth(int year, int month, int era = 0)
        {
            return _hebrewCalendar.IsLeapMonth(year, month, era);
        }

        public bool IsLeapYear(int year, int era = 0)
        {
            return _hebrewCalendar.IsLeapYear(year, era);
        }

        public DateTime ToDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond, int era = 0)
        {
            return new DateTime(_hebrewCalendar.ToDateTime(year, month, day, hour, minute, second, millisecond, era));
        }

        public int ToFourDigitYear(int year)
        {
            return _hebrewCalendar.ToFourDigitYear(year);
        }

        new public string ToString()
        { 
            return _hebrewCalendar.ToString(); 
        }


    }
}
