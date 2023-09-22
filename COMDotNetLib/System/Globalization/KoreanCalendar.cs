// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.koreancalendar?view=netframework-4.8.1

using GGlobalization = global::System.Globalization;
using GSystem = global::System;
using System;
using DotNetLib.Extensions;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Globalization;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("51C32295-CA96-4B46-9775-C836725153FA")]
    [ProgId("DotNetLib.System.Globalization.KoreanCalendar")]
    [Description("Represents the Korean calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IKoreanCalendar))]
    public class KoreanCalendar : IKoreanCalendar, Calendar, IWrappedObject
    {
        private GGlobalization.KoreanCalendar _koreanCalendar;

        //Constructors
        public KoreanCalendar()
        {
            _koreanCalendar = new GGlobalization.KoreanCalendar();
        }

        public KoreanCalendar(GGlobalization.KoreanCalendar koreanCalendar)
        {
            _koreanCalendar = koreanCalendar;
        }
        
        // Properties
        public GGlobalization.KoreanCalendar WrappedKoreanCalendar => _koreanCalendar;

        public object WrappedObject => _koreanCalendar;

        public CalendarAlgorithmType AlgorithmType => _koreanCalendar.AlgorithmType;

        public int[] Eras => _koreanCalendar.Eras;

        public bool IsReadOnly => _koreanCalendar.IsReadOnly;

        public DateTime MaxSupportedDateTime => new DateTime(_koreanCalendar.MaxSupportedDateTime);

        public DateTime MinSupportedDateTime => new DateTime(_koreanCalendar.MinSupportedDateTime);

        public int TwoDigitYearMax
        {
            get => _koreanCalendar.TwoDigitYearMax;
            set => _koreanCalendar.TwoDigitYearMax = value;
        }

        // Methods

        public DateTime AddDays(DateTime time, int days)
        {
            return new DateTime(_koreanCalendar.AddDays(time.WrappedDateTime, days));
        }

        public DateTime AddHours(DateTime time, int hours)
        {
            return new DateTime(_koreanCalendar.AddHours(time.WrappedDateTime, hours));
        }

        public DateTime AddMilliseconds(DateTime time, double milliseconds)
        {
            return new DateTime(_koreanCalendar.AddMilliseconds(time.WrappedDateTime, milliseconds));
        }

        public DateTime AddMinutes(DateTime time, int minutes)
        {
            return new DateTime(_koreanCalendar.AddMilliseconds(time.WrappedDateTime, minutes));
        }

        public DateTime AddMonths(DateTime time, int months)
        {
            return new DateTime(_koreanCalendar.AddMilliseconds(time.WrappedDateTime, months));
        }

        public DateTime AddSeconds(DateTime time, int seconds)
        {
            return new DateTime(_koreanCalendar.AddSeconds(time.WrappedDateTime, seconds));
        }

        public DateTime AddWeeks(DateTime time, int weeks)
        {
            return new DateTime(_koreanCalendar.AddWeeks(time.WrappedDateTime, weeks));
        }

        public DateTime AddYears(DateTime time, int years)
        {
            return new DateTime(_koreanCalendar.AddYears(time.WrappedDateTime, years));
        }

        public object Clone()
        {
            return new JapaneseCalendar((GGlobalization.JapaneseCalendar)_koreanCalendar.Clone());
        }

        public override bool Equals(object obj)
        {
            return obj is JapaneseCalendar cal && _koreanCalendar.Equals(cal.WrappedJapaneseCalendar);
        }

        public int GetDayOfMonth(DateTime time)
        {
            return _koreanCalendar.GetDayOfMonth(time.WrappedDateTime);
        }

        public GSystem.DayOfWeek GetDayOfWeek(DateTime time)
        {
            return _koreanCalendar.GetDayOfWeek(time.WrappedDateTime);
        }

        public int GetDayOfYear(DateTime time)
        {
            return _koreanCalendar.GetDayOfYear(time.WrappedDateTime);
        }

        public int GetDaysInMonth(int year, int month, int era = 0)
        {
            return _koreanCalendar.GetDaysInMonth(year, month, era);
        }

        public int GetDaysInYear(int year, int era = 0)
        {
            return _koreanCalendar.GetDaysInYear(year, era);
        }

        public int GetEra(DateTime time)
        {
            return _koreanCalendar.GetEra(time.WrappedDateTime);
        }

        public override int GetHashCode()
        {
            return _koreanCalendar.GetHashCode();
        }

        public int GetHour(DateTime time)
        {
            return _koreanCalendar.GetHour(time.WrappedDateTime);
        }

        public int GetLeapMonth(int year, int era = 0)
        {
            return _koreanCalendar.GetLeapMonth(year, era);
        }

        public double GetMilliseconds(DateTime time)
        {
            return _koreanCalendar.GetMilliseconds(time.WrappedDateTime);
        }

        public int GetMinute(DateTime time)
        {
            return _koreanCalendar.GetMinute(time.WrappedDateTime);
        }

        public int GetMonth(DateTime time)
        {
            return _koreanCalendar.GetMonth(time.WrappedDateTime);
        }

        public int GetMonthsInYear(int year, int era = 0)
        {
            return _koreanCalendar.GetMonthsInYear(year, era);
        }

        public int GetSecond(DateTime time)
        {
            return _koreanCalendar.GetSecond(time.WrappedDateTime);
        }

        public int GetWeekOfYear(DateTime time, CalendarWeekRule rule, GSystem.DayOfWeek firstDayOfWeek)
        {
            return _koreanCalendar.GetWeekOfYear(time.WrappedDateTime, rule, firstDayOfWeek);
        }

        public int GetYear(DateTime time)
        {
            return _koreanCalendar.GetYear(time.WrappedDateTime);
        }

        public bool IsLeapDay(int year, int month, int day, int era = 0)
        {
            return _koreanCalendar.IsLeapDay(year, month, day, era);
        }

        public bool IsLeapMonth(int year, int month, int era = 0)
        {
            return _koreanCalendar.IsLeapMonth(year, month, era);
        }

        public bool IsLeapYear(int year, int era = 0)
        {
            return _koreanCalendar.IsLeapYear(year, era);
        }

        public DateTime ToDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond, int era = 0)
        {
            return new DateTime(_koreanCalendar.ToDateTime(year, month, day, hour, minute, second, millisecond, era));
        }

        public int ToFourDigitYear(int year)
        {
            return _koreanCalendar.ToFourDigitYear(year);
        }

        new public string ToString()
        { 
            return _koreanCalendar.ToString();
        }

    }
}
