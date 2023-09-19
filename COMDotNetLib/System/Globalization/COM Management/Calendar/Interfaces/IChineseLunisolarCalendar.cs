//  https://learn.microsoft.com/en-us/dotnet/api/system.globalization.chineselunisolarcalendar?view=netframework-4.8.1

using GGlobalization = global::System.Globalization;
using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("79FF116E-6948-4673-A486-99D5EDEE3D7A")]
    [Description("Represents time in divisions, such as months, days, and years. Years are calculated using the Chinese calendar, while days and months are calculated using the lunisolar calendar.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IChineseLunisolarCalendar : ICalendar
    {
        //Properties

        new GGlobalization.CalendarAlgorithmType AlgorithmType
        {
            [Description("Gets a value indicating whether the current calendar is solar-based, lunar-based, or a combination of both.")]
            get;
        }

        new int[] Eras
        {
            [Description("Gets the eras that correspond to the range of dates and times supported by the current ChineseLunisolarCalendar object.")]
            get;
        }

        new bool IsReadOnly
        {
            [Description("Gets a value indicating whether this Calendar object is read-only.")]
            get;
        }

        new DateTime MaxSupportedDateTime
        {
            [Description("Gets the maximum date and time supported by the ChineseLunisolarCalendar class.")]
            get;
        }

        new DateTime MinSupportedDateTime
        {
            [Description("Gets the minimum date and time supported by the ChineseLunisolarCalendar class.")]
            get;
        }

        new int TwoDigitYearMax
        {
            [Description("Gets or sets the last year of a 100-year range that can be represented by a 2-digit year.")]
            get;
            [Description(" Gets or sets the last year of a 100-year range that can be represented by a 2-digit year.")]
            set;
        }

        //Methods

        [Description("Returns a DateTime that is the specified number of days away from the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        new DateTime AddDays(DateTime time, int days);

        [Description("Returns a DateTime that is the specified number of hours away from the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        new DateTime AddHours(DateTime time, int hours);

        [Description("Returns a DateTime that is the specified number of milliseconds away from the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        new DateTime AddMilliseconds(DateTime time, double milliseconds);

        [Description("Returns a DateTime that is the specified number of minutes away from the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        new DateTime AddMinutes(DateTime time, int minutes);

        [Description("Returns a DateTime that is the specified number of months away from the specified DateTime.")]
        new DateTime AddMonths(DateTime time, int months);

        [Description("Returns a DateTime that is the specified number of seconds away from the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        new DateTime AddSeconds(DateTime time, int seconds);

        [Description("Returns a DateTime that is the specified number of weeks away from the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        new DateTime AddWeeks(DateTime time, int weeks);

        [Description("Returns a DateTime that is the specified number of years away from the specified DateTime.")]
        new DateTime AddYears(DateTime time, int years);

        [Description("Creates a new object that is a copy of the current Calendar object.\r\n\r\n(Inherited from Calendar)")]
        new object Clone();

        [Description("Determines whether the specified object is equal to the current object.\r\n\r\n(Inherited from Object)")]
        new bool Equals(object obj);

        [Description("Returns the day of the month in the specified DateTime.")]
        new int GetDayOfMonth(DateTime time);

        [Description("Returns the day of the week in the specified DateTime.")]
        new GSystem.DayOfWeek GetDayOfWeek(DateTime time);

        [Description("Returns the day of the year in the specified DateTime.")]
        new int GetDayOfYear(DateTime time);

        [Description("Returns the number of days in the specified month, year, and era or the current era.")]
        new int GetDaysInMonth(int year, int month, int era = 0);

        [Description("Returns the number of days in the specified year in the specified era.")]
        new int GetDaysInYear(int year, int era = 0);

        [Description("Returns the era in the specified DateTime.")]
        new int GetEra(DateTime time);

        [Description("Serves as the default hash function.")]
        new int GetHashCode();

        [Description("Returns the hours value in the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        new int GetHour(DateTime time);

        [Description("Calculates the leap month for a specified year and era.")]
        new int GetLeapMonth(int year, int era = 0);

        [Description("Returns the milliseconds value in the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        new double GetMilliseconds(DateTime time);

        [Description("Returns the minutes value in the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        new int GetMinute(DateTime time);

        [Description("Returns the month in the specified DateTime.")]
        new int GetMonth(DateTime time);

        [Description("Returns the number of months in the specified year in the specified era.")]
        new int GetMonthsInYear(int year, int era = 0);

        [Description("Returns the seconds value in the specified DateTime.\r\n\r\n(Inherited from Calendar)")]
        new int GetSecond(DateTime time);

        [Description("Returns the week of the year that includes the date in the specified DateTime value.\r\n\r\n(Inherited from Calendar)")]
        new int GetWeekOfYear(DateTime time, GGlobalization.CalendarWeekRule rule, GSystem.DayOfWeek firstDayOfWeek);

        [Description("Returns the year in the specified DateTime.")]
        new int GetYear(DateTime time);

        [Description("Determines whether the specified date in the specified era is a leap day.")]
        new bool IsLeapDay(int year, int month, int day, int era = 0);

        [Description("Determines whether the specified month in the specified year in the specified era is a leap month.")]
        new bool IsLeapMonth(int year, int month, int era = 0);

        [Description("Determines whether the specified year in the specified era is a leap year.")]
        new bool IsLeapYear(int year, int era = 0);

        [Description("Returns a DateTime that is set to the specified date and time in the specified era.")]
        new DateTime ToDateTime(int year, int month, int day, int hour, int minute, int second, int millisecond, int era = 0);

        [Description("Converts the specified year to a four-digit year by using the TwoDigitYearMax property to determine the appropriate century.")]
        new int ToFourDigitYear(int year);

        [Description("Returns a string that represents the current object.\r\n\r\n(Inherited from Object)")]
        new string ToString();

        [Description("Calculates the celestial stem of the specified year in the sexagenary (60-year) cycle.")]
        int GetCelestialStem(int sexagenaryYear);

        [Description("Calculates the year in the sexagenary (60-year) cycle that corresponds to the specified date.")]
        int GetSexagenaryYear(DateTime time);

        [Description("Calculates the terrestrial branch of the specified year in the sexagenary (60-year) cycle.")]
        int GetTerrestrialBranch(int sexagenaryYear);

    }
}
