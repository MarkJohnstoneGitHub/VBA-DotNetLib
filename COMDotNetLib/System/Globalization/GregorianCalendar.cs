using GSystem = global::System;
using GGlobalization = global::System.Globalization;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Globalization;

namespace DotNetLib.System.Globalization
{
    [ComVisible(false)]
    [Guid("2F14BC67-B3F9-47BB-87C8-FEA1142120CC")]
    [ProgId("DotNetLib.System.Globalization.Calendar")]
    [Description("Represents the Gregorian calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IGregorianCalendar))]

    public class GregorianCalendar : GGlobalization.GregorianCalendar, ICalendar, IGregorianCalendar
    {

        //Constructors
        public GregorianCalendar(GGlobalization.GregorianCalendar gregorianCalendar)
        {
            GregorianCalendarBase = gregorianCalendar;
        }

        public GregorianCalendar() : base()
        {
            //GregorianCalendarBase = new GGlobalization.GregorianCalendar();
        }

        public GregorianCalendar(GregorianCalendarTypes type) : base(type)
        {
            //GregorianCalendarBase = new GGlobalization.GregorianCalendar(type);
        }

        //Fields
        public new int CurrentEra => CurrentEra;

        public new int DaysInYearBeforeMinSupportedYear => base.DaysInYearBeforeMinSupportedYear;

        public new DateTime MaxSupportedDateTime => new DateTime(base.MaxSupportedDateTime);

        public new DateTime MinSupportedDateTime => new DateTime(base.MinSupportedDateTime);

        public new int ADEra => GGlobalization.GregorianCalendar.ADEra;


        // Properties
        public GGlobalization.GregorianCalendar GregorianCalendarBase 
        { 
            get; 
            set; 
        }


        // Methods
        public DateTime AddDays(DateTime time, int days)
        {
            return new DateTime(base.AddDays(time.DateTimeObject,days));
        }

        public DateTime AddHours(DateTime time, int hours)
        {
            return new DateTime(base.AddHours(time.DateTimeObject, hours));
        }

        public DateTime AddMilliseconds(DateTime time, double milliseconds)
        {
            return new DateTime(base.AddMilliseconds(time.DateTimeObject, milliseconds));
        }

        public DateTime AddMinutes(DateTime time, int minutes)
        {
            return new DateTime(base.AddMilliseconds(time.DateTimeObject, minutes));
        }

        public DateTime AddMonths(DateTime time, int months)
        {
            return new DateTime(base.AddMilliseconds(time.DateTimeObject, months));
        }

        public DateTime AddSeconds(DateTime time, int seconds)
        {
            return new DateTime(base.AddSeconds(time.DateTimeObject, seconds));
        }

        public DateTime AddWeeks(DateTime time, int weeks)
        {
            return new DateTime(base.AddWeeks(time.DateTimeObject, weeks));
        }

        public DateTime AddYears(DateTime time, int years)
        {
            return new DateTime(base.AddYears(time.DateTimeObject, years));
        }

        public int GetDayOfMonth(DateTime time)
        {
            return base.GetDayOfMonth(time.DateTimeObject);
        }

        public GSystem.DayOfWeek GetDayOfWeek(DateTime time)
        {
            return base.GetDayOfWeek(time.DateTimeObject);
        }

        public int GetDayOfYear(DateTime time)
        {
            return base.GetDayOfYear(time.DateTimeObject);
        }

        public int GetDaysInMonth2(int year, int month, int era)
        {
            return base.GetDaysInMonth(year,month,era);
        }

        public int GetDaysInYear2(int year, int era)
        {
            return base.GetDaysInYear(year, era);
        }

        public int GetEra(DateTime time)
        {
            return base.GetEra(time.DateTimeObject);
        }

        public int GetHour(DateTime time)
        {
            return base.GetHour(time.DateTimeObject);
        }

        public int GetLeapMonth2(int year, int era)
        {
            return base.GetLeapMonth(year, era);
        }

        public double GetMilliseconds(DateTime time)
        {
            return base.GetMilliseconds(time.DateTimeObject);
        }

        public int GetMinute(DateTime time)
        {
            return base.GetMinute(time.DateTimeObject);
        }

        public int GetMonth(DateTime time)
        {
            return base.GetMonth(time.DateTimeObject);
        }

        public int GetMonthsInYear2(int year, int era)
        {
            return base.GetMonthsInYear(year, era);
        }

        public int GetSecond(DateTime time)
        {
            return base.GetSecond(time.DateTimeObject);
        }

        public int GetWeekOfYear(DateTime time, CalendarWeekRule rule, GSystem.DayOfWeek firstDayOfWeek)
        {
            return base.GetWeekOfYear(time.DateTimeObject,rule,firstDayOfWeek);
        }

        public int GetYear(DateTime time)
        {
            return base.GetYear(time.DateTimeObject);
        }

        public bool IsLeapDay2(int year, int month, int day, int era)
        {
            return base.IsLeapDay(year, month, day, era);
        }

        public bool IsLeapMonth2(int year, int month, int era)
        {
            return base.IsLeapMonth(year, month, era);
        }

        public bool IsLeapYear2(int year, int era)
        {
            return base.IsLeapYear(year, era);
        }

        public new DateTime ToDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond)
        {
            return new DateTime(base.ToDateTime(year, month, day, hour, minute, second, millisecond));
        }

        public DateTime ToDateTime2(int year, int month, int day, int hour, int minute, int second, int millisecond, int era)
        {
            return new DateTime(base.ToDateTime(year,month,day,hour,minute,second,millisecond,era));
        }

        public new object MemberwiseClone()
        {
            return new GregorianCalendar((GGlobalization.GregorianCalendar)base.MemberwiseClone());
        }
    }
}
