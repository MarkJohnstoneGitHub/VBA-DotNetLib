using System;
using System.Collections;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using GGlobalization = global::System.Globalization;
using GSystem = global::System;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("21E79313-3518-4798-9DD0-5978C9A72672")]
    [Description("Provides culture-specific information about the format of date and time values.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDateTimeFormatInfo
    {
        // Properties

        string[] AbbreviatedDayNames
        {
            [Description("Gets or sets a one-dimensional array of type String containing the culture-specific abbreviated names of the days of the week.")]
            get;
            //[Description("Gets or sets a one-dimensional array of type String containing the culture-specific abbreviated names of the days of the week.")]
            //set;
        }


        [Description("Sets a one-dimensional array of type String containing the culture-specific abbreviated names of the days of the week.")]
        void SetAbbreviatedDayNames([In] ref string[] abbreviatedDayNames);

        string[] AbbreviatedMonthGenitiveNames
        {
            [Description("Gets or sets a string array of abbreviated month names associated with the current DateTimeFormatInfo object.")]
            get;
            //[Description("Gets or sets a string array of abbreviated month names associated with the current DateTimeFormatInfo object.")]
            //set;
        }

        [Description("Sets a string array of abbreviated month names associated with the current DateTimeFormatInfo object.")]
        void SetAbbreviatedMonthGenitiveNames([In] ref string[] abbreviatedMonthGenitiveNames);

        string[] AbbreviatedMonthNames
        {
            [Description("Gets or sets a one-dimensional string array that contains the culture-specific abbreviated names of the months.")]
            get;
            //[Description("Gets or sets a one-dimensional string array that contains the culture-specific abbreviated names of the months.")]
            //set;
        }

        [Description("Sets a one-dimensional string array that contains the culture-specific abbreviated names of the months.")]
        void SetAbbreviatedMonthNames([In] ref string[] abbreviatedMonthNames);

        string AMDesignator 
        {
            [Description("Gets or sets the string designator for hours that are \"ante meridiem\" (before noon).")]
            get;
            [Description("Gets or sets the string designator for hours that are \"ante meridiem\" (before noon).")]
            set;
        }

        GGlobalization.Calendar Calendar 
        {
            [Description("Gets or sets the calendar to use for the current culture.")]
            get;
            [Description("Gets or sets the calendar to use for the current culture.")]
            set;
        }

        GGlobalization.CalendarWeekRule CalendarWeekRule 
        {
            [Description("Gets or sets a value that specifies which rule is used to determine the first calendar week of the year.")]
            get;
            [Description("Gets or sets a value that specifies which rule is used to determine the first calendar week of the year.")]
            set;
        }

        string DateSeparator 
        {
            [Description("Gets or sets the string that separates the components of a date, that is, the year, month, and day.")]
            get;
            [Description("Gets or sets the string that separates the components of a date, that is, the year, month, and day.")]
            set;
        }

        string[] DayNames 
        {
            [Description("Gets or sets a one-dimensional string array that contains the culture-specific full names of the days of the week.")]
            get;
            //[Description("Gets or sets a one-dimensional string array that contains the culture-specific full names of the days of the week.")]
            //set;
        }

        [Description("Sets a one-dimensional string array that contains the culture-specific full names of the days of the week.")]
        void SetDayNames([In] ref string[] dayNames);

        DayOfWeek FirstDayOfWeek 
        {
            [Description("Gets or sets the first day of the week.")]
            get;
            [Description("Gets or sets the first day of the week.")]
            set;
        }

        string FullDateTimePattern 
        {
            [Description("Gets or sets the custom format string for a long date and long time value.")]
            get;
            [Description("Gets or sets the custom format string for a long date and long time value.")]
            set;
        }

        bool IsReadOnly 
        {
            [Description("Gets a value indicating whether the DateTimeFormatInfo object is read-only.")]
            get;
        }

        string LongDatePattern 
        {
            [Description("Gets or sets the custom format string for a long date value.")]
            get;
            [Description("Gets or sets the custom format string for a long date value.")]
            set;
        }

        string LongTimePattern 
        {
            [Description("Gets or sets the custom format string for a long time value.")]
            get;
            [Description("Gets or sets the custom format string for a long time value.")]
            set;
        }

        string MonthDayPattern 
        {
            [Description("Gets or sets the custom format string for a month and day value.")]
            get;
            [Description("Gets or sets the custom format string for a month and day value.")]
            set;
        }

        string[] MonthGenitiveNames 
        {
            [Description("Gets or sets a string array of month names associated with the current DateTimeFormatInfo object.")]
            get;
            //[Description("Gets or sets a string array of month names associated with the current DateTimeFormatInfo object.")]
            //set;
        }

        [Description("Sets a string array of month names associated with the current DateTimeFormatInfo object.")]
        void SetMonthGenitiveNames([In] ref string[] monthGenitiveNames);


        string[] MonthNames 
        {
            [Description("Gets or sets a one-dimensional array of type String containing the culture-specific full names of the months.")]
            get;
            //[Description("Gets or sets a one-dimensional array of type String containing the culture-specific full names of the months.")]
            //set;
        }

        [Description("Sets a one-dimensional array of type String containing the culture-specific full names of the months.")]
        void SetMonthNames([In] ref string[] monthNames);

        string NativeCalendarName 
        {
            [Description("Gets the native name of the calendar associated with the current DateTimeFormatInfo object.")]
            get;
        }

        string PMDesignator 
        {
            [Description("Gets or sets the string designator for hours that are \"post meridiem\" (after noon).")]
            get;
            [Description("Gets or sets the string designator for hours that are \"post meridiem\" (after noon).")]
            set;
        }

        string RFC1123Pattern 
        {
            [Description("Gets the custom format string for a time value that is based on the Internet Engineering Task Force (IETF) Request for Comments (RFC) 1123 specification.")]
            get;
        }

        string ShortDatePattern 
        {
            [Description("Gets or sets the custom format string for a short date value.")]
            get;
            [Description("Gets or sets the custom format string for a short date value.")]
            set;
        }

        string[] ShortestDayNames 
        {
            [Description("Gets or sets a string array of the shortest unique abbreviated day names associated with the current DateTimeFormatInfo object.")]
            get;
            //[Description("Gets or sets a string array of the shortest unique abbreviated day names associated with the current DateTimeFormatInfo object.")]
            //set;
        }

        [Description("Sets a string array of the shortest unique abbreviated day names associated with the current DateTimeFormatInfo object.")]
        void SetShortestDayNames([In] ref string[] shortestDayNames);


        string ShortTimePattern 
        {
            [Description("Gets or sets the custom format string for a short time value.")]
            get;
            [Description("Gets or sets the custom format string for a short time value.")]
            set;
        }

        string SortableDateTimePattern 
        {
            [Description("Gets the custom format string for a sortable date and time value.")]
            get;
        }

        string TimeSeparator 
        {
            [Description("Gets or sets the string that separates the components of time, that is, the hour, minutes, and seconds.")]
            get;
            [Description("Gets or sets the string that separates the components of time, that is, the hour, minutes, and seconds.")]
            set;
        }

        string UniversalSortableDateTimePattern 
        {
            [Description("Gets the custom format string for a universal, sortable date and time string, as defined by ISO 8601.")]
            get;
        }

        string YearMonthPattern 
        {
            [Description("Gets or sets the custom format string for a year and month value.")]
            get;
            [Description("Gets or sets the custom format string for a year and month value.")]
            set;
        }

        // Methods

        [Description("Creates a shallow copy of the DateTimeFormatInfo.")]
        object Clone();

        [Description("Returns the culture-specific abbreviated name of the specified day of the week based on the culture associated with the current DateTimeFormatInfo object.")]
        string GetAbbreviatedDayName(DayOfWeek dayofweek);

        [Description("Returns the string containing the abbreviated name of the specified era, if an abbreviation exists.")]
        string GetAbbreviatedEraName(int era);

        [Description("Returns the string containing the abbreviated name of the specified era, if an abbreviation exists.")]
        string GetAbbreviatedMonthName(int month);

        [Description("Returns all the patterns in which date and time values can be formatted using the specified standard format string.")]
        string[] GetAllDateTimePatterns(string format = null);

        //[Description("Returns all the patterns in which date and time values can be formatted using the specified standard format string.")]
        //string[] GetAllDateTimePatterns2(char format);

        [Description("Returns the culture-specific full name of the specified day of the week based on the culture associated with the current DateTimeFormatInfo object.")]
        string GetDayName(DayOfWeek dayofweek);

        [Description("Returns the integer representing the specified era.")]
        int GetEra(string eraName);

        [Description("Returns the string containing the name of the specified era.")]
        string GetEraName(int era);

        [Description("Returns an object of the specified type that provides a date and time formatting service.")]
        object GetFormat(Type formatType);

        [Description("Returns the culture-specific full name of the specified month based on the culture associated with the current DateTimeFormatInfo object.")]
        string GetMonthName(int month);

        [Description("Obtains the shortest abbreviated day name for a specified day of the week associated with the current DateTimeFormatInfo object.")]
        string GetShortestDayName(DayOfWeek dayOfWeek);

        [Description("Sets the custom date and time format strings that correspond to a specified standard format string.")]
        void SetAllDateTimePatterns([In] ref string[] patterns, string format);

    }
}
