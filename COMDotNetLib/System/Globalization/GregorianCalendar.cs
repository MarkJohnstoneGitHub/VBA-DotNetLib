// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.gregoriancalendar?view=netframework-4.8.1

using GSystem = global::System;
using GGlobalization = global::System.Globalization;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Globalization;
using DotNetLib.Extensions;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("2F14BC67-B3F9-47BB-87C8-FEA1142120CC")]
    [ProgId("DotNetLib.System.Globalization.GregorianCalendar")]
    [Description("Represents the Gregorian calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IGregorianCalendar))]

    public class GregorianCalendar : IGregorianCalendar, Calendar, IWrappedObject
    {
        private GGlobalization.GregorianCalendar _gregorianCalendar;

        //Constructors
        public GregorianCalendar(GGlobalization.GregorianCalendar gregorianCalendar)
        {
            _gregorianCalendar = gregorianCalendar;
        }

        public GregorianCalendar() 
        {
            _gregorianCalendar = new GGlobalization.GregorianCalendar();
        }

        public GregorianCalendar(GregorianCalendarTypes type)
        {
            _gregorianCalendar = new GGlobalization.GregorianCalendar(type);
        }

        // Properties
        public GGlobalization.GregorianCalendar WrappedGregorianCalendar => _gregorianCalendar;

        public object WrappedObject => _gregorianCalendar;

        public CalendarAlgorithmType AlgorithmType => _gregorianCalendar.AlgorithmType;

        public GregorianCalendarTypes CalendarType 
        { 
            get => _gregorianCalendar.CalendarType; 
            set => _gregorianCalendar.CalendarType = value;
        }

        public int[] Eras => _gregorianCalendar.Eras;

        public bool IsReadOnly => _gregorianCalendar.IsReadOnly;

        public DateTime MaxSupportedDateTime => new DateTime(_gregorianCalendar.MaxSupportedDateTime);

        public DateTime MinSupportedDateTime => new DateTime(_gregorianCalendar.MinSupportedDateTime);

        public int TwoDigitYearMax 
        { 
            get => _gregorianCalendar.TwoDigitYearMax;
            set => _gregorianCalendar.TwoDigitYearMax = value;
        }

        // Methods
        public DateTime AddDays(DateTime time, int days)
        {
            return new DateTime(_gregorianCalendar.AddDays(time.WrappedDateTime,days));
        }

        public DateTime AddHours(DateTime time, int hours)
        {
            return new DateTime(_gregorianCalendar.AddHours(time.WrappedDateTime, hours));
        }

        public DateTime AddMilliseconds(DateTime time, double milliseconds)
        {
            return new DateTime(_gregorianCalendar.AddMilliseconds(time.WrappedDateTime, milliseconds));
        }

        public DateTime AddMinutes(DateTime time, int minutes)
        {
            return new DateTime(_gregorianCalendar.AddMilliseconds(time.WrappedDateTime, minutes));
        }

        public DateTime AddMonths(DateTime time, int months)
        {
            return new DateTime(_gregorianCalendar.AddMilliseconds(time.WrappedDateTime, months));
        }

        public DateTime AddSeconds(DateTime time, int seconds)
        {
            return new DateTime(_gregorianCalendar.AddSeconds(time.WrappedDateTime, seconds));
        }

        public DateTime AddWeeks(DateTime time, int weeks)
        {
            return new DateTime(_gregorianCalendar.AddWeeks(time.WrappedDateTime, weeks));
        }

        public DateTime AddYears(DateTime time, int years)
        {
            return new DateTime(_gregorianCalendar.AddYears(time.WrappedDateTime, years));
        }

        public object Clone()
        {
            return new GregorianCalendar((GGlobalization.GregorianCalendar)_gregorianCalendar.Clone());
        }

        public override bool Equals(object obj)
        {
            return obj is GregorianCalendar cal && _gregorianCalendar.Equals(cal.WrappedGregorianCalendar);
        }

        public int GetDayOfMonth(DateTime time)
        {
            return _gregorianCalendar.GetDayOfMonth(time.WrappedDateTime);
        }

        public GSystem.DayOfWeek GetDayOfWeek(DateTime time)
        {
            return _gregorianCalendar.GetDayOfWeek(time.WrappedDateTime);
        }

        public int GetDayOfYear(DateTime time)
        {
            return _gregorianCalendar.GetDayOfYear(time.WrappedDateTime);
        }

        public int GetDaysInMonth(int year, int month, int era = 0)
        {
            return _gregorianCalendar.GetDaysInMonth(year, month, era);
        }

        public int GetDaysInYear(int year, int era = 0)
        {
            return _gregorianCalendar.GetDaysInYear(year, era);
        }

        public int GetEra(DateTime time)
        {
            return _gregorianCalendar.GetEra(time.WrappedDateTime);
        }

        public override int GetHashCode()
        { 
            return _gregorianCalendar.GetHashCode(); 
        }

        public int GetHour(DateTime time)
        {
            return _gregorianCalendar.GetHour(time.WrappedDateTime);
        }

        public int GetLeapMonth(int year, int era = 0)
        {
            return _gregorianCalendar.GetLeapMonth(year, era);
        }

        public double GetMilliseconds(DateTime time)
        {
            return _gregorianCalendar.GetMilliseconds(time.WrappedDateTime);
        }

        public int GetMinute(DateTime time)
        {
            return _gregorianCalendar.GetMinute(time.WrappedDateTime);
        }

        public int GetMonth(DateTime time)
        {
            return _gregorianCalendar.GetMonth(time.WrappedDateTime);
        }

        public int GetMonthsInYear(int year, int era = 0)
        {
            return _gregorianCalendar.GetMonthsInYear(year, era);
        }

        public int GetSecond(DateTime time)
        {
            return _gregorianCalendar.GetSecond(time.WrappedDateTime);
        }

        public int GetWeekOfYear(DateTime time, CalendarWeekRule rule, GSystem.DayOfWeek firstDayOfWeek)
        {
            return _gregorianCalendar.GetWeekOfYear(time.WrappedDateTime,rule,firstDayOfWeek);
        }

        public int GetYear(DateTime time)
        {
            return _gregorianCalendar.GetYear(time.WrappedDateTime);
        }

        public bool IsLeapDay(int year, int month, int day, int era =  0)
        {
            return _gregorianCalendar.IsLeapDay(year, month, day, era);
        }

        public bool IsLeapMonth(int year, int month, int era = 0)
        {
            return _gregorianCalendar.IsLeapMonth(year, month, era);
        }

        public bool IsLeapYear(int year, int era = 0)
        {
            return _gregorianCalendar.IsLeapYear(year, era);
        }

        public DateTime ToDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond, int era = 0)
        {
            return new DateTime(_gregorianCalendar.ToDateTime(year,month,day,hour,minute,second,millisecond,era));
        }

        public int ToFourDigitYear(int year)
        {
            return _gregorianCalendar.ToFourDigitYear(year);
        }

        new public string ToString()
        { 
            return _gregorianCalendar.ToString(); 
        }
    }
}
