//  https://learn.microsoft.com/en-us/dotnet/api/system.globalization.chineselunisolarcalendar?view=netframework-4.8.1

using GSystem = global::System;
using GGlobalization = global::System.Globalization;
using DotNetLib.Extensions;
using System.Globalization;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("D06E0ECC-5B77-4217-A684-FE02911CE89C")]
    [ProgId("DotNetLib.System.Globalization.ChineseLunisolarCalendar")]
    [Description("Represents time in divisions, such as months, days, and years. Years are calculated using the Chinese calendar, while days and months are calculated using the lunisolar calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IChineseLunisolarCalendar))]
    public class ChineseLunisolarCalendar : IChineseLunisolarCalendar, ICalendar, IWrappedObject
    {
        private GGlobalization.ChineseLunisolarCalendar _chineseLunisolarCalendar;

        //Constructors
        public ChineseLunisolarCalendar()
        {
            _chineseLunisolarCalendar = new GGlobalization.ChineseLunisolarCalendar();
        }

        public ChineseLunisolarCalendar(GGlobalization.ChineseLunisolarCalendar chineseLunisolarCalendar)
        {
            _chineseLunisolarCalendar = chineseLunisolarCalendar;
        }

        // Properties
        public GGlobalization.ChineseLunisolarCalendar WrappedHebrewCalendar => _chineseLunisolarCalendar;

        public object WrappedObject => _chineseLunisolarCalendar;

        public CalendarAlgorithmType AlgorithmType => _chineseLunisolarCalendar.AlgorithmType;

        public int[] Eras => _chineseLunisolarCalendar.Eras;

        public bool IsReadOnly => _chineseLunisolarCalendar.IsReadOnly;

        public DateTime MaxSupportedDateTime => new DateTime(_chineseLunisolarCalendar.MaxSupportedDateTime);

        public DateTime MinSupportedDateTime => new DateTime(_chineseLunisolarCalendar.MinSupportedDateTime);

        public int TwoDigitYearMax
        {
            get => _chineseLunisolarCalendar.TwoDigitYearMax;
            set => _chineseLunisolarCalendar.TwoDigitYearMax = value;
        }

        // Methods

        public DateTime AddDays(DateTime time, int days)
        {
            return new DateTime(_chineseLunisolarCalendar.AddDays(time.WrappedDateTime, days));
        }

        public DateTime AddHours(DateTime time, int hours)
        {
            return new DateTime(_chineseLunisolarCalendar.AddHours(time.WrappedDateTime, hours));
        }

        public DateTime AddMilliseconds(DateTime time, double milliseconds)
        {
            return new DateTime(_chineseLunisolarCalendar.AddMilliseconds(time.WrappedDateTime, milliseconds));
        }

        public DateTime AddMinutes(DateTime time, int minutes)
        {
            return new DateTime(_chineseLunisolarCalendar.AddMilliseconds(time.WrappedDateTime, minutes));
        }

        public DateTime AddMonths(DateTime time, int months)
        {
            return new DateTime(_chineseLunisolarCalendar.AddMilliseconds(time.WrappedDateTime, months));
        }

        public DateTime AddSeconds(DateTime time, int seconds)
        {
            return new DateTime(_chineseLunisolarCalendar.AddSeconds(time.WrappedDateTime, seconds));
        }

        public DateTime AddWeeks(DateTime time, int weeks)
        {
            return new DateTime(_chineseLunisolarCalendar.AddWeeks(time.WrappedDateTime, weeks));
        }

        public DateTime AddYears(DateTime time, int years)
        {
            return new DateTime(_chineseLunisolarCalendar.AddYears(time.WrappedDateTime, years));
        }

        public object Clone()
        {
            return new HebrewCalendar((GGlobalization.HebrewCalendar)_chineseLunisolarCalendar.Clone());
        }

        public override bool Equals(object obj)
        {
            return obj is HebrewCalendar cal && _chineseLunisolarCalendar.Equals(cal.WrappedHebrewCalendar);
        }

        public int GetCelestialStem(int sexagenaryYear)
        {
            return _chineseLunisolarCalendar.GetCelestialStem(sexagenaryYear);
        }

        public int GetSexagenaryYear(DateTime time)
        {
            return _chineseLunisolarCalendar.GetSexagenaryYear(time.WrappedDateTime);
        }

        public int GetTerrestrialBranch(int sexagenaryYear)
        {
            return _chineseLunisolarCalendar.GetTerrestrialBranch(sexagenaryYear);
        }

        public int GetDayOfMonth(DateTime time)
        {
            return _chineseLunisolarCalendar.GetDayOfMonth(time.WrappedDateTime);
        }

        public GSystem.DayOfWeek GetDayOfWeek(DateTime time)
        {
            return _chineseLunisolarCalendar.GetDayOfWeek(time.WrappedDateTime);
        }

        public int GetDayOfYear(DateTime time)
        {
            return _chineseLunisolarCalendar.GetDayOfYear(time.WrappedDateTime);
        }

        public int GetDaysInMonth(int year, int month, int era = 0)
        {
            return _chineseLunisolarCalendar.GetDaysInMonth(year, month, era);
        }

        public int GetDaysInYear(int year, int era = 0)
        {
            return _chineseLunisolarCalendar.GetDaysInYear(year, era);
        }

        public int GetEra(DateTime time)
        {
            return _chineseLunisolarCalendar.GetEra(time.WrappedDateTime);
        }

        public override int GetHashCode()
        {
            return _chineseLunisolarCalendar.GetHashCode();
        }

        public int GetHour(DateTime time)
        {
            return _chineseLunisolarCalendar.GetHour(time.WrappedDateTime);
        }

        public int GetLeapMonth(int year, int era = 0)
        {
            return _chineseLunisolarCalendar.GetLeapMonth(year, era);
        }

        public double GetMilliseconds(DateTime time)
        {
            return _chineseLunisolarCalendar.GetMilliseconds(time.WrappedDateTime);
        }

        public int GetMinute(DateTime time)
        {
            return _chineseLunisolarCalendar.GetMinute(time.WrappedDateTime);
        }

        public int GetMonth(DateTime time)
        {
            return _chineseLunisolarCalendar.GetMonth(time.WrappedDateTime);
        }

        public int GetMonthsInYear(int year, int era = 0)
        {
            return _chineseLunisolarCalendar.GetMonthsInYear(year, era);
        }

        public int GetSecond(DateTime time)
        {
            return _chineseLunisolarCalendar.GetSecond(time.WrappedDateTime);
        }

        public int GetWeekOfYear(DateTime time, CalendarWeekRule rule, GSystem.DayOfWeek firstDayOfWeek)
        {
            return _chineseLunisolarCalendar.GetWeekOfYear(time.WrappedDateTime, rule, firstDayOfWeek);
        }

        public int GetYear(DateTime time)
        {
            return _chineseLunisolarCalendar.GetYear(time.WrappedDateTime);
        }

        public bool IsLeapDay(int year, int month, int day, int era = 0)
        {
            return _chineseLunisolarCalendar.IsLeapDay(year, month, day, era);
        }

        public bool IsLeapMonth(int year, int month, int era = 0)
        {
            return _chineseLunisolarCalendar.IsLeapMonth(year, month, era);
        }

        public bool IsLeapYear(int year, int era = 0)
        {
            return _chineseLunisolarCalendar.IsLeapYear(year, era);
        }

        public DateTime ToDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond, int era = 0)
        {
            return new DateTime(_chineseLunisolarCalendar.ToDateTime(year, month, day, hour, minute, second, millisecond, era));
        }

        public int ToFourDigitYear(int year)
        {
            return _chineseLunisolarCalendar.ToFourDigitYear(year);
        }

        new public string ToString()
        { 
            return _chineseLunisolarCalendar.ToString();
        }

    }
}
