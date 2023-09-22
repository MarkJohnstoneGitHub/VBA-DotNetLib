// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.thaibuddhistcalendar?view=netframework-4.8.1

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
    [Guid("ACF72703-D72D-47C5-9E23-6794355C3EA9")]
    [ProgId("DotNetLib.System.Globalization.ThaiBuddhistCalendar")]
    [Description("Represents the Thai Buddhist calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IThaiBuddhistCalendar))]
    public class ThaiBuddhistCalendar : IThaiBuddhistCalendar, Calendar, IWrappedObject
    {
        private GGlobalization.ThaiBuddhistCalendar _thaiBuddhistCalendar;

        //Constructors
        public ThaiBuddhistCalendar()
        {
            _thaiBuddhistCalendar = new GGlobalization.ThaiBuddhistCalendar();
        }

        public ThaiBuddhistCalendar(GGlobalization.ThaiBuddhistCalendar thaiBuddhistCalendar)
        {
            _thaiBuddhistCalendar = thaiBuddhistCalendar;
        }

        //Fields
        //public int CurrentEra => Calendar.CurrentEra;

        //public int ThaiBuddhistEra => GGlobalization.ThaiBuddhistCalendar.ThaiBuddhistEra;

        // Properties
        public GGlobalization.ThaiBuddhistCalendar WrappedThaiBuddhistCalendar => _thaiBuddhistCalendar;

        public object WrappedObject => _thaiBuddhistCalendar;

        public CalendarAlgorithmType AlgorithmType => _thaiBuddhistCalendar.AlgorithmType;

        public int[] Eras => _thaiBuddhistCalendar.Eras;

        public bool IsReadOnly => _thaiBuddhistCalendar.IsReadOnly;

        public DateTime MaxSupportedDateTime => new DateTime(_thaiBuddhistCalendar.MaxSupportedDateTime);

        public DateTime MinSupportedDateTime => new DateTime(_thaiBuddhistCalendar.MinSupportedDateTime);

        public int TwoDigitYearMax
        {
            get => _thaiBuddhistCalendar.TwoDigitYearMax;
            set => _thaiBuddhistCalendar.TwoDigitYearMax = value;
        }


        // Methods
        public DateTime AddDays(DateTime time, int days)
        {
            return new DateTime(_thaiBuddhistCalendar.AddDays(time.WrappedDateTime, days));
        }

        public DateTime AddHours(DateTime time, int hours)
        {
            return new DateTime(_thaiBuddhistCalendar.AddHours(time.WrappedDateTime, hours));
        }

        public DateTime AddMilliseconds(DateTime time, double milliseconds)
        {
            return new DateTime(_thaiBuddhistCalendar.AddMilliseconds(time.WrappedDateTime, milliseconds));
        }

        public DateTime AddMinutes(DateTime time, int minutes)
        {
            return new DateTime(_thaiBuddhistCalendar.AddMilliseconds(time.WrappedDateTime, minutes));
        }

        public DateTime AddMonths(DateTime time, int months)
        {
            return new DateTime(_thaiBuddhistCalendar.AddMilliseconds(time.WrappedDateTime, months));
        }

        public DateTime AddSeconds(DateTime time, int seconds)
        {
            return new DateTime(_thaiBuddhistCalendar.AddSeconds(time.WrappedDateTime, seconds));
        }

        public DateTime AddWeeks(DateTime time, int weeks)
        {
            return new DateTime(_thaiBuddhistCalendar.AddWeeks(time.WrappedDateTime, weeks));
        }

        public DateTime AddYears(DateTime time, int years)
        {
            return new DateTime(_thaiBuddhistCalendar.AddYears(time.WrappedDateTime, years));
        }

        public object Clone()
        {
            return new ThaiBuddhistCalendar((GGlobalization.ThaiBuddhistCalendar)_thaiBuddhistCalendar.Clone());
        }

        public override bool Equals(object obj)
        {
            return obj is ThaiBuddhistCalendar cal && _thaiBuddhistCalendar.Equals(cal.WrappedThaiBuddhistCalendar);
        }

        public int GetDayOfMonth(DateTime time)
        {
            return _thaiBuddhistCalendar.GetDayOfMonth(time.WrappedDateTime);
        }

        public GSystem.DayOfWeek GetDayOfWeek(DateTime time)
        {
            return _thaiBuddhistCalendar.GetDayOfWeek(time.WrappedDateTime);
        }

        public int GetDayOfYear(DateTime time)
        {
            return _thaiBuddhistCalendar.GetDayOfYear(time.WrappedDateTime);
        }

        public int GetDaysInMonth(int year, int month, int era = 0)
        {
            return _thaiBuddhistCalendar.GetDaysInMonth(year, month, era);
        }

        public int GetDaysInYear(int year, int era = 0)
        {
            return _thaiBuddhistCalendar.GetDaysInYear(year, era);
        }

        public int GetEra(DateTime time)
        {
            return _thaiBuddhistCalendar.GetEra(time.WrappedDateTime);
        }

        public override int GetHashCode()
        {
            return _thaiBuddhistCalendar.GetHashCode();
        }

        public int GetHour(DateTime time)
        {
            return _thaiBuddhistCalendar.GetHour(time.WrappedDateTime);
        }

        public int GetLeapMonth(int year, int era = 0)
        {
            return _thaiBuddhistCalendar.GetLeapMonth(year, era);
        }

        public double GetMilliseconds(DateTime time)
        {
            return _thaiBuddhistCalendar.GetMilliseconds(time.WrappedDateTime);
        }

        public int GetMinute(DateTime time)
        {
            return _thaiBuddhistCalendar.GetMinute(time.WrappedDateTime);
        }

        public int GetMonth(DateTime time)
        {
            return _thaiBuddhistCalendar.GetMonth(time.WrappedDateTime);
        }

        public int GetMonthsInYear(int year, int era = 0)
        {
            return _thaiBuddhistCalendar.GetMonthsInYear(year, era);
        }

        public int GetSecond(DateTime time)
        {
            return _thaiBuddhistCalendar.GetSecond(time.WrappedDateTime);
        }

        public int GetWeekOfYear(DateTime time, CalendarWeekRule rule, GSystem.DayOfWeek firstDayOfWeek)
        {
            return _thaiBuddhistCalendar.GetWeekOfYear(time.WrappedDateTime, rule, firstDayOfWeek);
        }

        public int GetYear(DateTime time)
        {
            return _thaiBuddhistCalendar.GetYear(time.WrappedDateTime);
        }

        public bool IsLeapDay(int year, int month, int day, int era = 0)
        {
            return _thaiBuddhistCalendar.IsLeapDay(year, month, day, era);
        }

        public bool IsLeapMonth(int year, int month, int era = 0)
        {
            return _thaiBuddhistCalendar.IsLeapMonth(year, month, era);
        }

        public bool IsLeapYear(int year, int era = 0)
        {
            return _thaiBuddhistCalendar.IsLeapYear(year, era);
        }

        public DateTime ToDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond, int era = 0)
        {
            return new DateTime(_thaiBuddhistCalendar.ToDateTime(year, month, day, hour, minute, second, millisecond, era));
        }

        public int ToFourDigitYear(int year)
        {
            return _thaiBuddhistCalendar.ToFourDigitYear(year);
        }

        new public string ToString()
        {
            return _thaiBuddhistCalendar.ToString();
        }

    }
}
