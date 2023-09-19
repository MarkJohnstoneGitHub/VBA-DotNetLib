// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.hijricalendar?view=netframework-4.8.1

using GSystem = global::System;
using GGlobalization = global::System.Globalization;
using DotNetLib.Extensions;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Globalization;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("00F730EB-EE2B-49F6-8513-C24E49F20202")]
    [ProgId("DotNetLib.System.Globalization.HijriCalendar")]
    [Description("Represents the Hijri calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IHijriCalendar))]
    public class HijriCalendar : IHijriCalendar, ICalendar, IWrappedObject
    {
        private GGlobalization.HijriCalendar _hijriCalendar;

        //Constructors
        public HijriCalendar()
        {
            _hijriCalendar = new GGlobalization.HijriCalendar();
        }

        public HijriCalendar(GGlobalization.HijriCalendar hijriCalendar)
        {
            _hijriCalendar = hijriCalendar;
        }

        // Properties
        public GGlobalization.HijriCalendar WrappedHijriCalendar => _hijriCalendar;

        public object WrappedObject => _hijriCalendar;

        public CalendarAlgorithmType AlgorithmType => _hijriCalendar.AlgorithmType;

        public int[] Eras => _hijriCalendar.Eras;

        public bool IsReadOnly => _hijriCalendar.IsReadOnly;

        public DateTime MaxSupportedDateTime => new DateTime(_hijriCalendar.MaxSupportedDateTime);

        public DateTime MinSupportedDateTime => new DateTime(_hijriCalendar.MinSupportedDateTime);

        public int TwoDigitYearMax
        {
            get => _hijriCalendar.TwoDigitYearMax;
            set => _hijriCalendar.TwoDigitYearMax = value;
        }

        // Methods

        public DateTime AddDays(DateTime time, int days)
        {
            return new DateTime(_hijriCalendar.AddDays(time.WrappedDateTime, days));
        }

        public DateTime AddHours(DateTime time, int hours)
        {
            return new DateTime(_hijriCalendar.AddHours(time.WrappedDateTime, hours));
        }

        public DateTime AddMilliseconds(DateTime time, double milliseconds)
        {
            return new DateTime(_hijriCalendar.AddMilliseconds(time.WrappedDateTime, milliseconds));
        }

        public DateTime AddMinutes(DateTime time, int minutes)
        {
            return new DateTime(_hijriCalendar.AddMilliseconds(time.WrappedDateTime, minutes));
        }

        public DateTime AddMonths(DateTime time, int months)
        {
            return new DateTime(_hijriCalendar.AddMilliseconds(time.WrappedDateTime, months));
        }

        public DateTime AddSeconds(DateTime time, int seconds)
        {
            return new DateTime(_hijriCalendar.AddSeconds(time.WrappedDateTime, seconds));
        }

        public DateTime AddWeeks(DateTime time, int weeks)
        {
            return new DateTime(_hijriCalendar.AddWeeks(time.WrappedDateTime, weeks));
        }

        public DateTime AddYears(DateTime time, int years)
        {
            return new DateTime(_hijriCalendar.AddYears(time.WrappedDateTime, years));
        }

        public object Clone()
        {
            return new HijriCalendar((GGlobalization.HijriCalendar)_hijriCalendar.Clone());
        }

        public override bool Equals(object obj)
        {
            return obj is HijriCalendar cal && _hijriCalendar.Equals(cal.WrappedHijriCalendar);
        }

        public int GetDayOfMonth(DateTime time)
        {
            return _hijriCalendar.GetDayOfMonth(time.WrappedDateTime);
        }

        public GSystem.DayOfWeek GetDayOfWeek(DateTime time)
        {
            return _hijriCalendar.GetDayOfWeek(time.WrappedDateTime);
        }

        public int GetDayOfYear(DateTime time)
        {
            return _hijriCalendar.GetDayOfYear(time.WrappedDateTime);
        }

        public int GetDaysInMonth(int year, int month, int era = 0)
        {
            return _hijriCalendar.GetDaysInMonth(year, month, era);
        }

        public int GetDaysInYear(int year, int era = 0)
        {
            return _hijriCalendar.GetDaysInYear(year, era);
        }

        public int GetEra(DateTime time)
        {
            return _hijriCalendar.GetEra(time.WrappedDateTime);
        }

        public override int GetHashCode()
        {
            return _hijriCalendar.GetHashCode();
        }

        public int GetHour(DateTime time)
        {
            return _hijriCalendar.GetHour(time.WrappedDateTime);
        }

        public int GetLeapMonth(int year, int era = 0)
        {
            return _hijriCalendar.GetLeapMonth(year, era);
        }

        public double GetMilliseconds(DateTime time)
        {
            return _hijriCalendar.GetMilliseconds(time.WrappedDateTime);
        }

        public int GetMinute(DateTime time)
        {
            return _hijriCalendar.GetMinute(time.WrappedDateTime);
        }

        public int GetMonth(DateTime time)
        {
            return _hijriCalendar.GetMonth(time.WrappedDateTime);
        }

        public int GetMonthsInYear(int year, int era = 0)
        {
            return _hijriCalendar.GetMonthsInYear(year, era);
        }

        public int GetSecond(DateTime time)
        {
            return _hijriCalendar.GetSecond(time.WrappedDateTime);
        }

        public int GetWeekOfYear(DateTime time, CalendarWeekRule rule, GSystem.DayOfWeek firstDayOfWeek)
        {
            return _hijriCalendar.GetWeekOfYear(time.WrappedDateTime, rule, firstDayOfWeek);
        }

        public int GetYear(DateTime time)
        {
            return _hijriCalendar.GetYear(time.WrappedDateTime);
        }

        public bool IsLeapDay(int year, int month, int day, int era = 0)
        {
            return _hijriCalendar.IsLeapDay(year, month, day, era);
        }

        public bool IsLeapMonth(int year, int month, int era = 0)
        {
            return _hijriCalendar.IsLeapMonth(year, month, era);
        }

        public bool IsLeapYear(int year, int era = 0)
        {
            return _hijriCalendar.IsLeapYear(year, era);
        }

        public DateTime ToDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond, int era = 0)
        {
            return new DateTime(_hijriCalendar.ToDateTime(year, month, day, hour, minute, second, millisecond, era));
        }

        public int ToFourDigitYear(int year)
        {
            return _hijriCalendar.ToFourDigitYear(year);
        }

        new public string ToString()
        { 
            return _hijriCalendar.ToString();
        }

    }
}
