// Notes : https://stackoverflow.com/questions/13185159/how-to-pass-byte-arrays-as-udt-properties-from-vb6-vba-to-c-sharp-com-dll

using GGlobalization = global::System.Globalization;
using System;
using System.Globalization;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("1448768E-E417-43D3-AA5C-8234F57324B9")]
    [ProgId("DotNetLib.System.DateTimeFormatInfo")]
    [Description("Provides culture-specific information about the format of date and time values.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDateTimeFormatInfo))]
    public class DateTimeFormatInfo : ICloneable, IFormatProvider, IDateTimeFormatInfo
    {
        private GGlobalization.DateTimeFormatInfo _dateTimeFormatInfo;

        public DateTimeFormatInfo()
        {
            _dateTimeFormatInfo = new GGlobalization.DateTimeFormatInfo();
        }

        public DateTimeFormatInfo(GGlobalization.DateTimeFormatInfo dateTimeFormatInfo)
        {
            _dateTimeFormatInfo = dateTimeFormatInfo;
        }

        public DateTimeFormatInfo(object v)
        {
            _dateTimeFormatInfo = (GGlobalization.DateTimeFormatInfo)v;
        }


        // Properties
        public GGlobalization.DateTimeFormatInfo DateTimeFormatInfoObject
        {
            get { return _dateTimeFormatInfo; }
            set { _dateTimeFormatInfo = value; }
        }

        public string[] AbbreviatedDayNames
        { 
            get => _dateTimeFormatInfo.AbbreviatedDayNames;
            set => _dateTimeFormatInfo.AbbreviatedDayNames = value;
        }
        public void SetAbbreviatedDayNames([In] ref string[] abbreviatedDayNames)
        {
            _dateTimeFormatInfo.AbbreviatedDayNames = abbreviatedDayNames;
        }

        public string[] AbbreviatedMonthGenitiveNames 
        { 
            get => _dateTimeFormatInfo.AbbreviatedMonthGenitiveNames;
            set => _dateTimeFormatInfo.AbbreviatedMonthGenitiveNames = value;
        }

        public void SetAbbreviatedMonthGenitiveNames([In] ref string[] abbreviatedMonthGenitiveNames)
        {
            _dateTimeFormatInfo.AbbreviatedMonthGenitiveNames = abbreviatedMonthGenitiveNames;
        }

        public string[] AbbreviatedMonthNames
        {
            get => _dateTimeFormatInfo.AbbreviatedMonthNames;
            set => _dateTimeFormatInfo.AbbreviatedMonthNames = value;
        }

        public void SetAbbreviatedMonthNames([In] ref string[] abbreviatedMonthNames)
        {
            _dateTimeFormatInfo.AbbreviatedMonthNames = abbreviatedMonthNames;
        }

        public string AMDesignator 
        { 
            get => _dateTimeFormatInfo.AMDesignator; 
            set => _dateTimeFormatInfo.AMDesignator = value;    
        }
        
        public Calendar Calendar 
        { 
            get => _dateTimeFormatInfo.Calendar;
            set => _dateTimeFormatInfo.Calendar = value;
        }

        public CalendarWeekRule CalendarWeekRule 
        { 
            get => _dateTimeFormatInfo.CalendarWeekRule;
            set => _dateTimeFormatInfo.CalendarWeekRule = value;
        }

        public static DateTimeFormatInfo CurrentInfo 
        {
            get => new DateTimeFormatInfo(GGlobalization.DateTimeFormatInfo.CurrentInfo);
        }

        public string DateSeparator 
        { 
            get => _dateTimeFormatInfo.DateSeparator;
            set => _dateTimeFormatInfo.DateSeparator = value;   
        }

        public string[] DayNames 
        { 
            get => _dateTimeFormatInfo.DayNames;
            set => _dateTimeFormatInfo.DayNames = value;
        }

        public void SetDayNames([In] ref string[] dayNames)
        {
            _dateTimeFormatInfo.DayNames = dayNames;
        }

        public DayOfWeek FirstDayOfWeek 
        { 
            get => (DayOfWeek)_dateTimeFormatInfo.FirstDayOfWeek;
            set => _dateTimeFormatInfo.FirstDayOfWeek = (global::System.DayOfWeek)value;
        }

        public string FullDateTimePattern 
        { 
            get => _dateTimeFormatInfo.FullDateTimePattern;
            set => _dateTimeFormatInfo.FullDateTimePattern = value;
        }

        public static DateTimeFormatInfo InvariantInfo
        {
            get => new DateTimeFormatInfo(GGlobalization.DateTimeFormatInfo.InvariantInfo);
        }

        public bool IsReadOnly => _dateTimeFormatInfo.IsReadOnly;

        public string LongDatePattern 
        { 
            get => _dateTimeFormatInfo.LongDatePattern;
            set => _dateTimeFormatInfo.LongDatePattern = value;
        }

        public string LongTimePattern 
        { 
            get => _dateTimeFormatInfo.LongTimePattern;
            set => _dateTimeFormatInfo.LongTimePattern = value;
        }

        public string MonthDayPattern 
        { 
            get => _dateTimeFormatInfo.MonthDayPattern;
            set => _dateTimeFormatInfo.MonthDayPattern = value;
        }

        public string[] MonthGenitiveNames 
        { 
            get => _dateTimeFormatInfo.MonthGenitiveNames; 
            set => _dateTimeFormatInfo.MonthGenitiveNames = value;
        }

        public void SetMonthGenitiveNames([In] ref string[] monthGenitiveNames)
        {
            _dateTimeFormatInfo.MonthGenitiveNames = monthGenitiveNames;
        }

        public string[] MonthNames 
        { 
            get => _dateTimeFormatInfo.MonthNames; 
            set => _dateTimeFormatInfo.MonthNames = value;
        }

        public void SetMonthNames([In] ref string[] monthNames)
        {
            _dateTimeFormatInfo.MonthNames = monthNames;
        }

        public string NativeCalendarName => _dateTimeFormatInfo.NativeCalendarName;

        public string PMDesignator 
        { 
            get => _dateTimeFormatInfo.PMDesignator; 
            set => _dateTimeFormatInfo.PMDesignator = value;
        }

        public string RFC1123Pattern => _dateTimeFormatInfo.RFC1123Pattern;

        public string ShortDatePattern 
        { 
            get => _dateTimeFormatInfo.ShortDatePattern;
            set => _dateTimeFormatInfo.ShortDatePattern = value;    
        }
        public string[] ShortestDayNames 
        { 
            get => _dateTimeFormatInfo.ShortestDayNames; 
            set => _dateTimeFormatInfo.ShortestDayNames = value;
        }

        public void SetShortestDayNames([In] ref string[] shortestDayNames)
        {
            _dateTimeFormatInfo.ShortestDayNames = ShortestDayNames;
        }

        public string ShortTimePattern 
        { 
            get => _dateTimeFormatInfo.ShortTimePattern;
            set => _dateTimeFormatInfo.ShortTimePattern = value;
        }

        public string SortableDateTimePattern => _dateTimeFormatInfo.SortableDateTimePattern;

        public string TimeSeparator 
        { 
            get => _dateTimeFormatInfo.TimeSeparator;
            set => _dateTimeFormatInfo.TimeSeparator = value;
        }

        public string UniversalSortableDateTimePattern => _dateTimeFormatInfo.UniversalSortableDateTimePattern;

        public string YearMonthPattern 
        { 
            get => _dateTimeFormatInfo.YearMonthPattern;
            set => _dateTimeFormatInfo.YearMonthPattern = value;
        }

        // Methods
        public object Clone()
        {
            return new NumberFormatInfo(_dateTimeFormatInfo.Clone);
        }

        public string GetAbbreviatedDayName(DayOfWeek dayofweek)
        {
            return _dateTimeFormatInfo.GetAbbreviatedDayName((global::System.DayOfWeek)dayofweek);
        }

        public string GetAbbreviatedEraName(int era)
        {
            return _dateTimeFormatInfo.GetAbbreviatedEraName(era);
        }

        public string GetAbbreviatedMonthName(int month)
        {
            return _dateTimeFormatInfo.GetAbbreviatedMonthName(month);
        }

        //public string[] GetAllDateTimePatterns()
        //{
        //    return _dateTimeFormatInfo.GetAllDateTimePatterns();
        //}

        public string[] GetAllDateTimePatterns(string format = null)
        {
            if (format == null)
            {
                return _dateTimeFormatInfo.GetAllDateTimePatterns();
            }
            if (format.Length > 1)
            {
                throw new ArgumentException("Format specifier was invalid.", "format");
            }
            return _dateTimeFormatInfo.GetAllDateTimePatterns(format[0]);
        }

        public string GetDayName(DayOfWeek dayofweek)
        {
            return _dateTimeFormatInfo.GetDayName((global::System.DayOfWeek)dayofweek);
        }

        public int GetEra(string eraName)
        {
            return _dateTimeFormatInfo.GetEra(eraName);
        }

        public string GetEraName(int era)
        {
            return _dateTimeFormatInfo.GetEraName(era);
        }

        // TODO: Check implementation
        public object GetFormat(Type formatType)
        {
            if (!(formatType == typeof(DateTimeFormatInfo)))
            {
                return null;
            }
            return this;
        }

        //TODO: Could also cater for mscorlib.DateTimeFormatInfo
        /// <summary>
        /// Unboxes the wrapped mscorlib.DateTimeFormatInfo
        /// </summary>
        /// <param name="provider"></param>
        /// <returns></returns>
        internal static GGlobalization.DateTimeFormatInfo GetFormatProvider(IFormatProvider provider)
        {
            CultureInfo cultureInfo = provider as CultureInfo;
            if (cultureInfo != null)
            {
                return cultureInfo.DateTimeFormat.DateTimeFormatInfoObject;
            }

            DateTimeFormatInfo dateTimeFormatInfo = provider as DateTimeFormatInfo;
            if (dateTimeFormatInfo != null)
            {
                return dateTimeFormatInfo.DateTimeFormatInfoObject;
            }

            if (provider != null)
            {
                dateTimeFormatInfo = provider.GetFormat(typeof(DateTimeFormatInfo)) as DateTimeFormatInfo;
                if (dateTimeFormatInfo != null)
                {
                    return dateTimeFormatInfo.DateTimeFormatInfoObject;
                }
            }

            return CurrentInfo.DateTimeFormatInfoObject;
        }

        public static DateTimeFormatInfo GetInstance(IFormatProvider formatProvider)
        {
            
            return new DateTimeFormatInfo(GGlobalization.DateTimeFormatInfo.GetInstance(GetFormatProvider(formatProvider)));
        }

        public string GetMonthName(int month)
        {
            return _dateTimeFormatInfo.GetMonthName(month);
        }

        public string GetShortestDayName(DayOfWeek dayOfWeek)
        {
            return _dateTimeFormatInfo.GetShortestDayName((global::System.DayOfWeek)dayOfWeek);
        }

        public static DateTimeFormatInfo ReadOnly(DateTimeFormatInfo dtfi)
        {
            return new DateTimeFormatInfo(GGlobalization.DateTimeFormatInfo.ReadOnly(dtfi.DateTimeFormatInfoObject));
        }

        public void SetAllDateTimePatterns([In] ref string[] patterns, string format)
        {
            if (format.Length == 0 | format.Length > 1 )
                {
                throw new ArgumentException("Format specifier was invalid.", "format");
            }
            _dateTimeFormatInfo.SetAllDateTimePatterns(patterns, format[0]);
        }

    }
}
