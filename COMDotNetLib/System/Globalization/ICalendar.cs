// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.calendar?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using GGlobalization = global::System.Globalization;
using GSystem = global::System;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("7CDC6E2F-4D2E-4053-9031-9F97A10260CF")]
    [Description("Represents time in divisions, such as weeks, months, and years.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICalendar
    {

        //Properties
        GGlobalization.CalendarAlgorithmType AlgorithmType 
        {
            [Description("Gets a value indicating whether the current calendar is solar-based, lunar-based, or a combination of both.")]
            get; 
        }

        //int DaysInYearBeforeMinSupportedYear 
        //{
        //    [Description("Gets the number of days in the year that precedes the year that is specified by the MinSupportedDateTime property.")]
        //    get;
        //}

        int[] Eras 
        {
            [Description("When overridden in a derived class, gets the list of eras in the current calendar.")]
            get; 
        }

        bool IsReadOnly 
        {
            [Description("Gets a value indicating whether this Calendar object is read-only.")]
            get;
        }

        DateTime MaxSupportedDateTime 
        {
            [Description("Gets the latest date and time supported by this Calendar object.")]
            get; 
        }

        DateTime MinSupportedDateTime 
        {
            [Description("Gets the earliest date and time supported by this Calendar object.")]
            get; 
        }

        int TwoDigitYearMax 
        {
            [Description("Gets or sets the last year of a 100-year range that can be represented by a 2-digit year.")]
            get;
            [Description(" Gets or sets the last year of a 100-year range that can be represented by a 2-digit year.")]
            set;
        }

        //Methods

        [Description("Returns a DateTime that is the specified number of days away from the specified DateTime.")]
        DateTime AddDays(DateTime time, int days);

        [Description("Returns a DateTime that is the specified number of hours away from the specified DateTime.")]
        DateTime AddHours(DateTime time, int hours);

        [Description("Returns a DateTime that is the specified number of milliseconds away from the specified DateTime.")]
        DateTime AddMilliseconds(DateTime time, double milliseconds);

        [Description("Returns a DateTime that is the specified number of minutes away from the specified DateTime.")]
        DateTime AddMinutes(DateTime time, int minutes);

        [Description("When overridden in a derived class, returns a DateTime that is the specified number of months away from the specified DateTime.")]
        DateTime AddMonths(DateTime time, int months);

        [Description("Returns a DateTime that is the specified number of seconds away from the specified DateTime.")]
        DateTime AddSeconds(DateTime time, int seconds);

        [Description("Returns a DateTime that is the specified number of weeks away from the specified DateTime.")]
        DateTime AddWeeks(DateTime time, int weeks);

        [Description("When overridden in a derived class, returns a DateTime that is the specified number of years away from the specified DateTime.")]
        DateTime AddYears(DateTime time, int years);

        [Description("Creates a new object that is a copy of the current Calendar object.")]
        object Clone();

        [Description("Determines whether the specified object is equal to the current object.\r\n\r\n(Inherited from Object)")]
        bool Equals(object obj);

        [Description("When overridden in a derived class, returns the day of the month in the specified DateTime.")]
        int GetDayOfMonth(DateTime time);

        [Description("When overridden in a derived class, returns the day of the week in the specified DateTime.")]
        GSystem.DayOfWeek GetDayOfWeek(DateTime time);

        [Description("")]
        int GetDayOfYear(DateTime time);

        [Description("Returns the number of days in the specified month, year, and era or the current era.")]
        int GetDaysInMonth(int year, int month, int era = 0);

        [Description("When overridden in a derived class, returns the number of days in the specified year and era.")]
        int GetDaysInYear(int year, int era = 0);

        [Description("When overridden in a derived class, returns the era of the specified DateTime.")]
        int GetEra(DateTime time);

        [Description("Serves as the default hash function.\r\n\r\n(Inherited from Object)")]
        int GetHashCode();


        [Description("Returns the hours value in the specified DateTime.")]
        int GetHour(DateTime time);

        [Description("Calculates the leap month for a specified year and era.")]
        int GetLeapMonth(int year, int era = 0);

        [Description("Returns the milliseconds value in the specified DateTime.")]
        double GetMilliseconds(DateTime time);

        [Description("Returns the minutes value in the specified DateTime.")]
        int GetMinute(DateTime time);

        [Description("When overridden in a derived class, returns the month in the specified DateTime.")]
        int GetMonth(DateTime time);

        [Description("When overridden in a derived class, returns the number of months in the specified year in the specified era.")]
        int GetMonthsInYear(int year, int era = 0);

        [Description("Returns the seconds value in the specified DateTime.")]
        int GetSecond(DateTime time);

        [Description("Returns the week of the year that includes the date in the specified DateTime value.")]
        int GetWeekOfYear(DateTime time, GGlobalization.CalendarWeekRule rule, GSystem.DayOfWeek firstDayOfWeek);

        [Description("When overridden in a derived class, returns the year in the specified DateTime.")]
        int GetYear(DateTime time);

        [Description("When overridden in a derived class, determines whether the specified date in the specified era is a leap day.")]
        bool IsLeapDay(int year, int month, int day, int era = 0);

        [Description("When overridden in a derived class, determines whether the specified month in the specified year in the specified era is a leap month.")]
        bool IsLeapMonth(int year, int month, int era = 0);

        [Description("When overridden in a derived class, determines whether the specified year in the specified era is a leap year.")]
        bool IsLeapYear(int year, int era = 0);

        [Description("When overridden in a derived class, returns a DateTime that is set to the specified date and time in the specified era.")]
        DateTime ToDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond, int era = 0);

        [Description("Converts the specified year to a four-digit year by using the TwoDigitYearMax property to determine the appropriate century.")]
        int ToFourDigitYear(int year);

        [Description("Returns a string that represents the current object.\r\n\r\n(Inherited from Object)")]
        string ToString();
    }
}
