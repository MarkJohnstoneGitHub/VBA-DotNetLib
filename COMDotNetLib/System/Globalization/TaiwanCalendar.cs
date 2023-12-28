// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.taiwancalendar?view=netframework-4.8.1

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
    [Guid("481F508E-5129-4751-BD21-53C399319C7C")]
    [ProgId("DotNetLib.System.Globalization.TaiwanCalendar")]
    [Description("The Taiwan calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITaiwanCalendar))]
    public class TaiwanCalendar : ITaiwanCalendar, Calendar, IWrappedObject
    {
        private GGlobalization.TaiwanCalendar _taiwanCalendar;

        //Constructors
        public TaiwanCalendar(GGlobalization.TaiwanCalendar taiwanCalendar)
        {
            _taiwanCalendar = taiwanCalendar;
        }

        public TaiwanCalendar()
        {
            _taiwanCalendar = new GGlobalization.TaiwanCalendar();
        }

        // Properties
        public GGlobalization.TaiwanCalendar WrappedTaiwanCalendar => _taiwanCalendar;

        public object WrappedObject => _taiwanCalendar;

        public CalendarAlgorithmType AlgorithmType => _taiwanCalendar.AlgorithmType;

        public int[] Eras => _taiwanCalendar.Eras;

        public bool IsReadOnly => _taiwanCalendar.IsReadOnly;

        public DateTime MaxSupportedDateTime => new DateTime(_taiwanCalendar.MaxSupportedDateTime);

        public DateTime MinSupportedDateTime => new DateTime(_taiwanCalendar.MinSupportedDateTime);

        public int TwoDigitYearMax
        {
            get => _taiwanCalendar.TwoDigitYearMax;
            set => _taiwanCalendar.TwoDigitYearMax = value;
        }

        // Methods
        public DateTime AddDays(DateTime time, int days)
        {
            return new DateTime(_taiwanCalendar.AddDays(time.WrappedDateTime, days));
        }

        public DateTime AddHours(DateTime time, int hours)
        {
            return new DateTime(_taiwanCalendar.AddHours(time.WrappedDateTime, hours));
        }

        public DateTime AddMilliseconds(DateTime time, double milliseconds)
        {
            return new DateTime(_taiwanCalendar.AddMilliseconds(time.WrappedDateTime, milliseconds));
        }

        public DateTime AddMinutes(DateTime time, int minutes)
        {
            return new DateTime(_taiwanCalendar.AddMilliseconds(time.WrappedDateTime, minutes));
        }

        public DateTime AddMonths(DateTime time, int months)
        {
            return new DateTime(_taiwanCalendar.AddMilliseconds(time.WrappedDateTime, months));
        }

        public DateTime AddSeconds(DateTime time, int seconds)
        {
            return new DateTime(_taiwanCalendar.AddSeconds(time.WrappedDateTime, seconds));
        }

        public DateTime AddWeeks(DateTime time, int weeks)
        {
            return new DateTime(_taiwanCalendar.AddWeeks(time.WrappedDateTime, weeks));
        }

        public DateTime AddYears(DateTime time, int years)
        {
            return new DateTime(_taiwanCalendar.AddYears(time.WrappedDateTime, years));
        }

        public object Clone()
        {
            return new TaiwanCalendar((GGlobalization.TaiwanCalendar)_taiwanCalendar.Clone());
        }

        public override bool Equals(object obj)
        {
            return obj is TaiwanCalendar cal && _taiwanCalendar.Equals(cal.WrappedTaiwanCalendar);
        }

        public int GetDayOfMonth(DateTime time)
        {
            return _taiwanCalendar.GetDayOfMonth(time.WrappedDateTime);
        }

        public GSystem.DayOfWeek GetDayOfWeek(DateTime time)
        {
            return _taiwanCalendar.GetDayOfWeek(time.WrappedDateTime);
        }

        public int GetDayOfYear(DateTime time)
        {
            return _taiwanCalendar.GetDayOfYear(time.WrappedDateTime);
        }

        public int GetDaysInMonth(int year, int month, int era = 0)
        {
            return _taiwanCalendar.GetDaysInMonth(year, month, era);
        }

        public int GetDaysInYear(int year, int era = 0)
        {
            return _taiwanCalendar.GetDaysInYear(year, era);
        }

        public int GetEra(DateTime time)
        {
            return _taiwanCalendar.GetEra(time.WrappedDateTime);
        }

        public override int GetHashCode()
        {
            return _taiwanCalendar.GetHashCode();
        }

        public int GetHour(DateTime time)
        {
            return _taiwanCalendar.GetHour(time.WrappedDateTime);
        }

        public int GetLeapMonth(int year, int era = 0)
        {
            return _taiwanCalendar.GetLeapMonth(year, era);
        }

        public double GetMilliseconds(DateTime time)
        {
            return _taiwanCalendar.GetMilliseconds(time.WrappedDateTime);
        }

        public int GetMinute(DateTime time)
        {
            return _taiwanCalendar.GetMinute(time.WrappedDateTime);
        }

        public int GetMonth(DateTime time)
        {
            return _taiwanCalendar.GetMonth(time.WrappedDateTime);
        }

        public int GetMonthsInYear(int year, int era = 0)
        {
            return _taiwanCalendar.GetMonthsInYear(year, era);
        }

        public int GetSecond(DateTime time)
        {
            return _taiwanCalendar.GetSecond(time.WrappedDateTime);
        }

        public int GetWeekOfYear(DateTime time, CalendarWeekRule rule, GSystem.DayOfWeek firstDayOfWeek)
        {
            return _taiwanCalendar.GetWeekOfYear(time.WrappedDateTime, rule, firstDayOfWeek);
        }

        public int GetYear(DateTime time)
        {
            return _taiwanCalendar.GetYear(time.WrappedDateTime);
        }

        public bool IsLeapDay(int year, int month, int day, int era = 0)
        {
            return _taiwanCalendar.IsLeapDay(year, month, day, era);
        }

        public bool IsLeapMonth(int year, int month, int era = 0)
        {
            return _taiwanCalendar.IsLeapMonth(year, month, era);
        }

        public bool IsLeapYear(int year, int era = 0)
        {
            return _taiwanCalendar.IsLeapYear(year, era);
        }

        public DateTime ToDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond, int era = 0)
        {
            return new DateTime(_taiwanCalendar.ToDateTime(year, month, day, hour, minute, second, millisecond, era));
        }

        public int ToFourDigitYear(int year)
        {
            return _taiwanCalendar.ToFourDigitYear(year);
        }

        new public string ToString()
        { 
            return _taiwanCalendar.ToString();
        }

    }
}
