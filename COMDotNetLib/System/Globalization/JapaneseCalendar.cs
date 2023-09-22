// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.japanesecalendar?view=netframework-4.8.1

using GSystem = global::System;
using GGlobalization = global::System.Globalization;
using DotNetLib.Extensions;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{

    [ComVisible(true)]
    [Guid("7C5B4C65-C72A-42E6-B60F-126E0B1A75E3")]
    [ProgId("DotNetLib.System.Globalization.JapaneseCalendar")]
    [Description("Represents the Japanese calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IJapaneseCalendar))]
    public class JapaneseCalendar : IJapaneseCalendar, Calendar, IWrappedObject
    {
        private GGlobalization.JapaneseCalendar _japaneseCalendar;

        //Constructors
        public JapaneseCalendar()
        {
            _japaneseCalendar = new GGlobalization.JapaneseCalendar();
        }

        public JapaneseCalendar(GGlobalization.JapaneseCalendar japaneseCalendar)
        {
            _japaneseCalendar = japaneseCalendar;
        }

        // Properties
        public GGlobalization.JapaneseCalendar WrappedJapaneseCalendar => _japaneseCalendar;

        public object WrappedObject => _japaneseCalendar;

        public CalendarAlgorithmType AlgorithmType => _japaneseCalendar.AlgorithmType;

        public int[] Eras => _japaneseCalendar.Eras;

        public bool IsReadOnly => _japaneseCalendar.IsReadOnly;

        public DateTime MaxSupportedDateTime => new DateTime(_japaneseCalendar.MaxSupportedDateTime);

        public DateTime MinSupportedDateTime => new DateTime(_japaneseCalendar.MinSupportedDateTime);

        public int TwoDigitYearMax
        {
            get => _japaneseCalendar.TwoDigitYearMax;
            set => _japaneseCalendar.TwoDigitYearMax = value;
        }

        // Methods

        public DateTime AddDays(DateTime time, int days)
        {
            return new DateTime(_japaneseCalendar.AddDays(time.WrappedDateTime, days));
        }

        public DateTime AddHours(DateTime time, int hours)
        {
            return new DateTime(_japaneseCalendar.AddHours(time.WrappedDateTime, hours));
        }

        public DateTime AddMilliseconds(DateTime time, double milliseconds)
        {
            return new DateTime(_japaneseCalendar.AddMilliseconds(time.WrappedDateTime, milliseconds));
        }

        public DateTime AddMinutes(DateTime time, int minutes)
        {
            return new DateTime(_japaneseCalendar.AddMilliseconds(time.WrappedDateTime, minutes));
        }

        public DateTime AddMonths(DateTime time, int months)
        {
            return new DateTime(_japaneseCalendar.AddMilliseconds(time.WrappedDateTime, months));
        }

        public DateTime AddSeconds(DateTime time, int seconds)
        {
            return new DateTime(_japaneseCalendar.AddSeconds(time.WrappedDateTime, seconds));
        }

        public DateTime AddWeeks(DateTime time, int weeks)
        {
            return new DateTime(_japaneseCalendar.AddWeeks(time.WrappedDateTime, weeks));
        }

        public DateTime AddYears(DateTime time, int years)
        {
            return new DateTime(_japaneseCalendar.AddYears(time.WrappedDateTime, years));
        }

        public object Clone()
        {
            return new JapaneseCalendar((GGlobalization.JapaneseCalendar)_japaneseCalendar.Clone());
        }

        public override bool Equals(object obj)
        {
            return obj is JapaneseCalendar cal && _japaneseCalendar.Equals(cal.WrappedJapaneseCalendar);
        }

        public int GetDayOfMonth(DateTime time)
        {
            return _japaneseCalendar.GetDayOfMonth(time.WrappedDateTime);
        }

        public GSystem.DayOfWeek GetDayOfWeek(DateTime time)
        {
            return _japaneseCalendar.GetDayOfWeek(time.WrappedDateTime);
        }

        public int GetDayOfYear(DateTime time)
        {
            return _japaneseCalendar.GetDayOfYear(time.WrappedDateTime);
        }

        public int GetDaysInMonth(int year, int month, int era = 0)
        {
            return _japaneseCalendar.GetDaysInMonth(year, month, era);
        }

        public int GetDaysInYear(int year, int era = 0)
        {
            return _japaneseCalendar.GetDaysInYear(year, era);
        }

        public int GetEra(DateTime time)
        {
            return _japaneseCalendar.GetEra(time.WrappedDateTime);
        }

        public override int GetHashCode()
        {
            return _japaneseCalendar.GetHashCode();
        }

        public int GetHour(DateTime time)
        {
            return _japaneseCalendar.GetHour(time.WrappedDateTime);
        }

        public int GetLeapMonth(int year, int era = 0)
        {
            return _japaneseCalendar.GetLeapMonth(year, era);
        }

        public double GetMilliseconds(DateTime time)
        {
            return _japaneseCalendar.GetMilliseconds(time.WrappedDateTime);
        }

        public int GetMinute(DateTime time)
        {
            return _japaneseCalendar.GetMinute(time.WrappedDateTime);
        }

        public int GetMonth(DateTime time)
        {
            return _japaneseCalendar.GetMonth(time.WrappedDateTime);
        }

        public int GetMonthsInYear(int year, int era = 0)
        {
            return _japaneseCalendar.GetMonthsInYear(year, era);
        }

        public int GetSecond(DateTime time)
        {
            return _japaneseCalendar.GetSecond(time.WrappedDateTime);
        }

        public int GetWeekOfYear(DateTime time, CalendarWeekRule rule, GSystem.DayOfWeek firstDayOfWeek)
        {
            return _japaneseCalendar.GetWeekOfYear(time.WrappedDateTime, rule, firstDayOfWeek);
        }

        public int GetYear(DateTime time)
        {
            return _japaneseCalendar.GetYear(time.WrappedDateTime);
        }

        public bool IsLeapDay(int year, int month, int day, int era = 0)
        {
            return _japaneseCalendar.IsLeapDay(year, month, day, era);
        }

        public bool IsLeapMonth(int year, int month, int era = 0)
        {
            return _japaneseCalendar.IsLeapMonth(year, month, era);
        }

        public bool IsLeapYear(int year, int era = 0)
        {
            return _japaneseCalendar.IsLeapYear(year, era);
        }

        public DateTime ToDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond, int era = 0)
        {
            return new DateTime(_japaneseCalendar.ToDateTime(year, month, day, hour, minute, second, millisecond, era));
        }

        public int ToFourDigitYear(int year)
        {
            return _japaneseCalendar.ToFourDigitYear(year);
        }

        new public string ToString()
        { 
            return _japaneseCalendar.ToString(); 
        }    

    }
}
