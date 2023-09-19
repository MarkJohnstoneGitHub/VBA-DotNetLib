// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.umalquracalendar?view=netframework-4.8.1

using GSystem = global::System;
using GGlobalization = global::System.Globalization;
using System.Globalization;
using DotNetLib.Extensions;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("D640BC74-261D-4C25-87F2-5322BC47EDF8")]
    [ProgId("DotNetLib.System.Globalization.UmAlQuraCalendar")]
    [Description("Represents the Saudi Hijri (Um Al Qura) calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IUmAlQuraCalendar))]
    public class UmAlQuraCalendar : IUmAlQuraCalendar, ICalendar, IWrappedObject
    {

        private GGlobalization.UmAlQuraCalendar _umAlQuraCalendar;

        //Constructors
        public UmAlQuraCalendar()
        {
            _umAlQuraCalendar = new GGlobalization.UmAlQuraCalendar();
        }

        public UmAlQuraCalendar(GGlobalization.UmAlQuraCalendar umAlQuraCalendar)
        {
            _umAlQuraCalendar = umAlQuraCalendar;
        }

        //Fields
        //public int CurrentEra => Calendar.CurrentEra;

        //public int UmAlQuraEra => GGlobalization.UmAlQuraCalendar.UmAlQuraEra;

        // Properties
        public GGlobalization.UmAlQuraCalendar WrappedUmAlQuraCalendar => _umAlQuraCalendar;

        public object WrappedObject => _umAlQuraCalendar;

        public CalendarAlgorithmType AlgorithmType => _umAlQuraCalendar.AlgorithmType;

        public int[] Eras => _umAlQuraCalendar.Eras;

        public bool IsReadOnly => _umAlQuraCalendar.IsReadOnly;

        public DateTime MaxSupportedDateTime => new DateTime(_umAlQuraCalendar.MaxSupportedDateTime);

        public DateTime MinSupportedDateTime => new DateTime(_umAlQuraCalendar.MinSupportedDateTime);

        public int TwoDigitYearMax
        {
            get => _umAlQuraCalendar.TwoDigitYearMax;
            set => _umAlQuraCalendar.TwoDigitYearMax = value;
        }


        // Methods
        public DateTime AddDays(DateTime time, int days)
        {
            return new DateTime(_umAlQuraCalendar.AddDays(time.WrappedDateTime, days));
        }

        public DateTime AddHours(DateTime time, int hours)
        {
            return new DateTime(_umAlQuraCalendar.AddHours(time.WrappedDateTime, hours));
        }

        public DateTime AddMilliseconds(DateTime time, double milliseconds)
        {
            return new DateTime(_umAlQuraCalendar.AddMilliseconds(time.WrappedDateTime, milliseconds));
        }

        public DateTime AddMinutes(DateTime time, int minutes)
        {
            return new DateTime(_umAlQuraCalendar.AddMilliseconds(time.WrappedDateTime, minutes));
        }

        public DateTime AddMonths(DateTime time, int months)
        {
            return new DateTime(_umAlQuraCalendar.AddMilliseconds(time.WrappedDateTime, months));
        }

        public DateTime AddSeconds(DateTime time, int seconds)
        {
            return new DateTime(_umAlQuraCalendar.AddSeconds(time.WrappedDateTime, seconds));
        }

        public DateTime AddWeeks(DateTime time, int weeks)
        {
            return new DateTime(_umAlQuraCalendar.AddWeeks(time.WrappedDateTime, weeks));
        }

        public DateTime AddYears(DateTime time, int years)
        {
            return new DateTime(_umAlQuraCalendar.AddYears(time.WrappedDateTime, years));
        }

        public object Clone()
        {
            return new UmAlQuraCalendar((GGlobalization.UmAlQuraCalendar)_umAlQuraCalendar.Clone());
        }

        public override bool Equals(object obj)
        {
            return obj is UmAlQuraCalendar cal && _umAlQuraCalendar.Equals(cal.WrappedUmAlQuraCalendar);
        }

        public int GetDayOfMonth(DateTime time)
        {
            return _umAlQuraCalendar.GetDayOfMonth(time.WrappedDateTime);
        }

        public GSystem.DayOfWeek GetDayOfWeek(DateTime time)
        {
            return _umAlQuraCalendar.GetDayOfWeek(time.WrappedDateTime);
        }

        public int GetDayOfYear(DateTime time)
        {
            return _umAlQuraCalendar.GetDayOfYear(time.WrappedDateTime);
        }

        public int GetDaysInMonth(int year, int month, int era = 0)
        {
            return _umAlQuraCalendar.GetDaysInMonth(year, month, era);
        }

        public int GetDaysInYear(int year, int era = 0)
        {
            return _umAlQuraCalendar.GetDaysInYear(year, era);
        }

        public int GetEra(DateTime time)
        {
            return _umAlQuraCalendar.GetEra(time.WrappedDateTime);
        }

        public override int GetHashCode()
        {
            return _umAlQuraCalendar.GetHashCode();
        }

        public int GetHour(DateTime time)
        {
            return _umAlQuraCalendar.GetHour(time.WrappedDateTime);
        }

        public int GetLeapMonth(int year, int era = 0)
        {
            return _umAlQuraCalendar.GetLeapMonth(year, era);
        }

        public double GetMilliseconds(DateTime time)
        {
            return _umAlQuraCalendar.GetMilliseconds(time.WrappedDateTime);
        }

        public int GetMinute(DateTime time)
        {
            return _umAlQuraCalendar.GetMinute(time.WrappedDateTime);
        }

        public int GetMonth(DateTime time)
        {
            return _umAlQuraCalendar.GetMonth(time.WrappedDateTime);
        }

        public int GetMonthsInYear(int year, int era = 0)
        {
            return _umAlQuraCalendar.GetMonthsInYear(year, era);
        }

        public int GetSecond(DateTime time)
        {
            return _umAlQuraCalendar.GetSecond(time.WrappedDateTime);
        }

        public int GetWeekOfYear(DateTime time, CalendarWeekRule rule, GSystem.DayOfWeek firstDayOfWeek)
        {
            return _umAlQuraCalendar.GetWeekOfYear(time.WrappedDateTime, rule, firstDayOfWeek);
        }

        public int GetYear(DateTime time)
        {
            return _umAlQuraCalendar.GetYear(time.WrappedDateTime);
        }

        public bool IsLeapDay(int year, int month, int day, int era = 0)
        {
            return _umAlQuraCalendar.IsLeapDay(year, month, day, era);
        }

        public bool IsLeapMonth(int year, int month, int era = 0)
        {
            return _umAlQuraCalendar.IsLeapMonth(year, month, era);
        }

        public bool IsLeapYear(int year, int era = 0)
        {
            return _umAlQuraCalendar.IsLeapYear(year, era);
        }

        public DateTime ToDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond, int era = 0)
        {
            return new DateTime(_umAlQuraCalendar.ToDateTime(year, month, day, hour, minute, second, millisecond, era));
        }

        public int ToFourDigitYear(int year)
        {
            return _umAlQuraCalendar.ToFourDigitYear(year);
        }

        new public string ToString()
        {
            return _umAlQuraCalendar.ToString();
        }
    }
}
