// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.persiancalendar?view=netframework-4.8.1

using GSystem = global::System;
using GGlobalization = global::System.Globalization;
using System.Globalization;
using DotNetLib.Extensions;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("F65E6BF1-7DDA-43C7-97F1-062CC0E1C954")]
    [ProgId("DotNetLib.System.Globalization.PersianCalendar")]
    [Description("Represents the Persian calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IPersianCalendar))]
    public class PersianCalendar : IPersianCalendar, Calendar, IWrappedObject
    {
        private GGlobalization.PersianCalendar _persianCalendar;

        //Constructors
        public PersianCalendar()
        {
            _persianCalendar = new GGlobalization.PersianCalendar();
        }

        public PersianCalendar(GGlobalization.PersianCalendar persianCalendar)
        {
            _persianCalendar = persianCalendar;
        }

        // Properties
        public GGlobalization.PersianCalendar WrappedPersianCalendar => _persianCalendar;

        public object WrappedObject => _persianCalendar;

        public CalendarAlgorithmType AlgorithmType => _persianCalendar.AlgorithmType;

        public int[] Eras => _persianCalendar.Eras;

        public bool IsReadOnly => _persianCalendar.IsReadOnly;

        public DateTime MaxSupportedDateTime => new DateTime(_persianCalendar.MaxSupportedDateTime);

        public DateTime MinSupportedDateTime => new DateTime(_persianCalendar.MinSupportedDateTime);

        public int TwoDigitYearMax
        {
            get => _persianCalendar.TwoDigitYearMax;
            set => _persianCalendar.TwoDigitYearMax = value;
        }


        // Methods
        public DateTime AddDays(DateTime time, int days)
        {
            return new DateTime(_persianCalendar.AddDays(time.WrappedDateTime, days));
        }

        public DateTime AddHours(DateTime time, int hours)
        {
            return new DateTime(_persianCalendar.AddHours(time.WrappedDateTime, hours));
        }

        public DateTime AddMilliseconds(DateTime time, double milliseconds)
        {
            return new DateTime(_persianCalendar.AddMilliseconds(time.WrappedDateTime, milliseconds));
        }

        public DateTime AddMinutes(DateTime time, int minutes)
        {
            return new DateTime(_persianCalendar.AddMilliseconds(time.WrappedDateTime, minutes));
        }

        public DateTime AddMonths(DateTime time, int months)
        {
            return new DateTime(_persianCalendar.AddMilliseconds(time.WrappedDateTime, months));
        }

        public DateTime AddSeconds(DateTime time, int seconds)
        {
            return new DateTime(_persianCalendar.AddSeconds(time.WrappedDateTime, seconds));
        }

        public DateTime AddWeeks(DateTime time, int weeks)
        {
            return new DateTime(_persianCalendar.AddWeeks(time.WrappedDateTime, weeks));
        }

        public DateTime AddYears(DateTime time, int years)
        {
            return new DateTime(_persianCalendar.AddYears(time.WrappedDateTime, years));
        }

        public object Clone()
        {
            return new PersianCalendar((GGlobalization.PersianCalendar)_persianCalendar.Clone());
        }

        public override bool Equals(object obj)
        {
            return obj is PersianCalendar cal && _persianCalendar.Equals(cal.WrappedPersianCalendar);
        }

        public int GetDayOfMonth(DateTime time)
        {
            return _persianCalendar.GetDayOfMonth(time.WrappedDateTime);
        }

        public GSystem.DayOfWeek GetDayOfWeek(DateTime time)
        {
            return _persianCalendar.GetDayOfWeek(time.WrappedDateTime);
        }

        public int GetDayOfYear(DateTime time)
        {
            return _persianCalendar.GetDayOfYear(time.WrappedDateTime);
        }

        public int GetDaysInMonth(int year, int month, int era = 0)
        {
            return _persianCalendar.GetDaysInMonth(year, month, era);
        }

        public int GetDaysInYear(int year, int era = 0)
        {
            return _persianCalendar.GetDaysInYear(year, era);
        }

        public int GetEra(DateTime time)
        {
            return _persianCalendar.GetEra(time.WrappedDateTime);
        }

        public override int GetHashCode()
        {
            return _persianCalendar.GetHashCode();
        }

        public int GetHour(DateTime time)
        {
            return _persianCalendar.GetHour(time.WrappedDateTime);
        }

        public int GetLeapMonth(int year, int era = 0)
        {
            return _persianCalendar.GetLeapMonth(year, era);
        }

        public double GetMilliseconds(DateTime time)
        {
            return _persianCalendar.GetMilliseconds(time.WrappedDateTime);
        }

        public int GetMinute(DateTime time)
        {
            return _persianCalendar.GetMinute(time.WrappedDateTime);
        }

        public int GetMonth(DateTime time)
        {
            return _persianCalendar.GetMonth(time.WrappedDateTime);
        }

        public int GetMonthsInYear(int year, int era = 0)
        {
            return _persianCalendar.GetMonthsInYear(year, era);
        }

        public int GetSecond(DateTime time)
        {
            return _persianCalendar.GetSecond(time.WrappedDateTime);
        }

        public int GetWeekOfYear(DateTime time, CalendarWeekRule rule, GSystem.DayOfWeek firstDayOfWeek)
        {
            return _persianCalendar.GetWeekOfYear(time.WrappedDateTime, rule, firstDayOfWeek);
        }

        public int GetYear(DateTime time)
        {
            return _persianCalendar.GetYear(time.WrappedDateTime);
        }

        public bool IsLeapDay(int year, int month, int day, int era = 0)
        {
            return _persianCalendar.IsLeapDay(year, month, day, era);
        }

        public bool IsLeapMonth(int year, int month, int era = 0)
        {
            return _persianCalendar.IsLeapMonth(year, month, era);
        }

        public bool IsLeapYear(int year, int era = 0)
        {
            return _persianCalendar.IsLeapYear(year, era);
        }

        public DateTime ToDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond, int era = 0)
        {
            return new DateTime(_persianCalendar.ToDateTime(year, month, day, hour, minute, second, millisecond, era));
        }

        public int ToFourDigitYear(int year)
        {
            return _persianCalendar.ToFourDigitYear(year);
        }

        new public string ToString()
        { 
            return _persianCalendar.ToString(); 
        } 

    }
}
