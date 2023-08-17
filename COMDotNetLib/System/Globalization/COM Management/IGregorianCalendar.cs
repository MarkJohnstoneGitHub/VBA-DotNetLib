using GGlobalization = global::System.Globalization;
using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;


namespace DotNetLib.System.Globalization
{
    [ComVisible(false)]
    [Guid("6B301F54-31F4-4263-BCB4-41F386441682")]
    [Description("Represents the Gregorian calendar.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IGregorianCalendar
    {
        //Fields

        int ADEra
        {
            [Description("Represents the current era. This field is constant.")]
            get;
        }

        int CurrentEra
        {
            [Description("Represents the current era of the current calendar. The value of this field is 0.")]
            get;
        }

        //Properties

        GGlobalization.GregorianCalendar GregorianCalendarBase
        {
            [Description("Gets or sets the GregorianCalendar base object")]
            get;
            [Description("Gets or sets the GregorianCalendar base object")]
            set;
        }

        GGlobalization.CalendarAlgorithmType AlgorithmType
        {
            [Description("Gets a value indicating whether the current calendar is solar-based, lunar-based, or a combination of both.")]
            get;
        }

        GGlobalization.GregorianCalendarTypes CalendarType
        {
            [Description("Gets or sets the GregorianCalendarTypes value that denotes the language version of the current GregorianCalendar.")]
            get;
            [Description("Gets or sets the GregorianCalendarTypes value that denotes the language version of the current GregorianCalendar.")]
            set;
        }

        int DaysInYearBeforeMinSupportedYear
        {
            [Description("Gets the number of days in the year that precedes the year that is specified by the MinSupportedDateTime property.\r\n\r\n(Inherited from Calendar)")]
            get;
        }
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

        [Description("Returns a DateTime that is the specified number of days away from the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        DateTime AddDays(DateTime time, int days);

        [Description("Returns a DateTime that is the specified number of hours away from the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        DateTime AddHours(DateTime time, int hours);

        [Description("Returns a DateTime that is the specified number of milliseconds away from the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        DateTime AddMilliseconds(DateTime time, double milliseconds);

        [Description("Returns a DateTime that is the specified number of minutes away from the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        DateTime AddMinutes(DateTime time, int minutes);

        [Description("Returns a DateTime that is the specified number of months away from the specified DateTime.")]
        DateTime AddMonths(DateTime time, int months);

        [Description("Returns a DateTime that is the specified number of seconds away from the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        DateTime AddSeconds(DateTime time, int seconds);

        [Description("Returns a DateTime that is the specified number of weeks away from the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        DateTime AddWeeks(DateTime time, int weeks);

        [Description("Returns a DateTime that is the specified number of years away from the specified DateTime.")]
        DateTime AddYears(DateTime time, int years);

        [Description("Creates a new object that is a copy of the current Calendar object.\r\n\r\n(Inherited from Calendar)")]
        object Clone();

        [Description("Determines whether the specified object is equal to the current object.\r\n\r\n(Inherited from Object)")]
        bool Equals(object obj);

        [Description("Returns the day of the month in the specified DateTime.")]
        int GetDayOfMonth(DateTime time);

        [Description("Returns the day of the week in the specified DateTime.")]
        GSystem.DayOfWeek GetDayOfWeek(DateTime time);

        [Description("Returns the day of the year in the specified DateTime.")]
        int GetDayOfYear(DateTime time);

        [Description("Returns the number of days in the specified month and year of the current era.\r\n\r\n(Inherited from Calendar)")]
        int GetDaysInMonth(int year, int month);

        [Description("Returns the number of days in the specified month in the specified year in the specified era.")]
        int GetDaysInMonth2(int year, int month, int era);

        [Description("Returns the number of days in the specified year of the current era.\r\n\r\n(Inherited from Calendar)")]
        int GetDaysInYear(int year);

        [Description("Returns the number of days in the specified year in the specified era.")]
        int GetDaysInYear2(int year, int era);

        [Description("Returns the era in the specified DateTime.")]
        int GetEra(DateTime time);

        [Description("Serves as the default hash function.")]
        int GetHashCode();

        [Description("Returns the hours value in the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        int GetHour(DateTime time);

        [Description("Calculates the leap month for a specified year.\r\n\r\n(Inherited from Calendar)")]
        int GetLeapMonth(int year);

        [Description("Calculates the leap month for a specified year and era.")]
        int GetLeapMonth2(int year, int era);

        [Description("Returns the milliseconds value in the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        double GetMilliseconds(DateTime time);

        [Description("Returns the minutes value in the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        int GetMinute(DateTime time);

        [Description("Returns the month in the specified DateTime.")]
        int GetMonth(DateTime time);

        [Description("Returns the number of months in the specified year in the current era.\r\n\r\n(Inherited from Calendar)")]
        int GetMonthsInYear(int year);

        [Description("Returns the number of months in the specified year in the specified era.")]
        int GetMonthsInYear2(int year, int era);

        [Description("Returns the seconds value in the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        int GetSecond(DateTime time);

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

        [Description("Returns the week of the year that includes the date in the specified DateTime value.\r\n\r\n(Inherited from Calendar)")]
        int GetWeekOfYear(DateTime time, GGlobalization.CalendarWeekRule rule, GSystem.DayOfWeek firstDayOfWeek);

        [Description("Returns the year in the specified DateTime.")]
        int GetYear(DateTime time);

        [Description("Determines whether the specified date in the current era is a leap day.\r\n\r\n(Inherited from Calendar)")]
        bool IsLeapDay(int year, int month, int day);

        [Description("Determines whether the specified date in the specified era is a leap day.")]
        bool IsLeapDay2(int year, int month, int day, int era);

        [Description("Determines whether the specified month in the specified year in the current era is a leap month.\r\n\r\n(Inherited from Calendar)")]
        bool IsLeapMonth(int year, int month);

        [Description("Determines whether the specified month in the specified year in the specified era is a leap month.")]
        bool IsLeapMonth2(int year, int month, int era);

        [Description("Determines whether the specified year in the current era is a leap year.\r\n\r\n(Inherited from Calendar)")]
        bool IsLeapYear(int year);

        [Description("Determines whether the specified year in the specified era is a leap year.")]
        bool IsLeapYear2(int year, int era);

        [Description("Creates a shallow copy of the current Object.\r\n\r\n(Inherited from Object)")]
        object MemberwiseClone();

        [Description("Returns a DateTime that is set to the specified date and time in the current era.\r\n\r\n(Inherited from Calendar)")]
        DateTime ToDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond);

        [Description("Returns a DateTime that is set to the specified date and time in the specified era.")]
        DateTime ToDateTime2(int year, int month, int day, int hour, int minute, int second, int millisecond, int era);

        [Description("Converts the specified year to a four-digit year by using the TwoDigitYearMax property to determine the appropriate century.")]
        int ToFourDigitYear(int year);

        [Description("Returns a string that represents the current object.\r\n\r\n(Inherited from Object)")]
        string ToString();
    }
}
