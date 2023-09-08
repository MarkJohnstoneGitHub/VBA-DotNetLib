// https://learn.microsoft.com/en-us/dotnet/api/system.timespan?view=netframework-4.8
// https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/TimeSpan.cs

using GSystem = global::System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System;
using DotNetLib.System.Globalization;
using DotNetLib.Extensions;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("B73DFD69-6C69-4CFC-89F2-1C344228A9D4")]
    [ProgId("DotNetLib.System.TimeSpan")]
    [Description("Represents a time interval.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITimeSpan))]
    public class TimeSpan : ITimeSpan, IWrappedObject 
    {
        private GSystem.TimeSpan _timeSpan;

        // Static Fields
        private static readonly TimeSpan tsMaxValue = new TimeSpan(GSystem.TimeSpan.MaxValue);
        private static readonly TimeSpan tsMinValue = new TimeSpan(GSystem.TimeSpan.MinValue);
        private static readonly TimeSpan tsZero = new TimeSpan(GSystem.TimeSpan.Zero);

        //Constructors
        public TimeSpan()
        {
            _timeSpan = new GSystem.TimeSpan();
        }

        internal TimeSpan(GSystem.TimeSpan value)
        {
            _timeSpan = value;
        }

        internal TimeSpan(long ticks)
        {
            _timeSpan = new GSystem.TimeSpan(ticks);
        }

        internal TimeSpan(int hours, int minutes, int seconds)
        {
            _timeSpan = new GSystem.TimeSpan(hours, minutes, seconds);
        }

        internal TimeSpan(int days, int hours, int minutes, int seconds)
        {
            _timeSpan = new GSystem.TimeSpan(days, hours, minutes, seconds);
        }

        internal TimeSpan(int days, int hours, int minutes, int seconds, int milliseconds)
        {
            _timeSpan = new GSystem.TimeSpan(days, hours, minutes, seconds, milliseconds);   
        }

        // Fields

        static public TimeSpan MaxValue => tsMaxValue;
        
        static public TimeSpan MinValue => tsMinValue;

        static public long TicksPerDay => GSystem.TimeSpan.TicksPerDay;
        
        static public long TicksPerHour => GSystem.TimeSpan.TicksPerHour;

        static public long TicksPerMillisecond => GSystem.TimeSpan.TicksPerMillisecond;

        static public long TicksPerMinute => GSystem.TimeSpan.TicksPerMinute;

        static public long TicksPerSecond => GSystem.TimeSpan.TicksPerSecond;

        static public TimeSpan Zero => tsZero;


        //Properties
        internal GSystem.TimeSpan WrappedTimeSpan
        {
            get { return _timeSpan; }
            set { _timeSpan = value; } 
        }

        public object WrappedObject => _timeSpan;

        public int Days => _timeSpan.Days;

        public int Hours => _timeSpan.Hours;

        public int Milliseconds => _timeSpan.Milliseconds;

        public int Minutes => _timeSpan.Minutes;

        public int Seconds => _timeSpan.Seconds;

        public long Ticks => _timeSpan.Ticks;

        public double TotalDays => _timeSpan.TotalDays;

        public double TotalHours => _timeSpan.TotalHours;

        public double TotalMilliseconds => _timeSpan.TotalMilliseconds;
        
        public double TotalMinutes => _timeSpan.TotalMinutes;
        
        public double TotalSeconds => _timeSpan.TotalSeconds;


        //Methods
        public TimeSpan Add(TimeSpan ts)
        {
            return new TimeSpan(_timeSpan.Add(ts.WrappedTimeSpan));
        }

        static public int Compare(TimeSpan ts1, TimeSpan ts2)
        {
            return GSystem.TimeSpan.Compare(ts1.WrappedTimeSpan,ts2.WrappedTimeSpan);  
        }

        public int CompareTo(object value)
        {
            return _timeSpan.CompareTo(value.Unwrap());

            //const string Arg_MustBeTimeSpan = "Object must be of type TimeSpan.";

            //if (value == null) return 1;
            //if (!(value is TimeSpan ts))
            //{
            //    throw new ArgumentException(Arg_MustBeTimeSpan);
            //}
            //return _timeSpan.CompareTo(ts.WrappedTimeSpan);
        }

        public int CompareTo2(TimeSpan value)
        {
            return _timeSpan.CompareTo(value.WrappedTimeSpan);
        }

        public TimeSpan Duration()
        {
            return new TimeSpan(_timeSpan.Duration());
        }

        public bool Equals(TimeSpan obj)
        { 
            return _timeSpan.Equals(obj.WrappedTimeSpan); 
        }

        // TODO : Check Implementation
        public bool Equals2(object value)
        {
            return value is TimeSpan ts && _timeSpan == ts.WrappedTimeSpan;
        }

        public static bool Equals(TimeSpan t1, TimeSpan t2)
        { 
            return GSystem.TimeSpan.Equals(t1.WrappedTimeSpan, t2.WrappedTimeSpan); 
        }

        static public TimeSpan FromDays(double value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromDays(value));
        }

        static public TimeSpan FromHours(double value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromHours(value));
        }

        static public TimeSpan FromMilliseconds(double value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromMilliseconds(value));
        }

        static public TimeSpan FromMinutes(double value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromMinutes(value));
        }

        static public TimeSpan FromSeconds(double value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromSeconds(value));
        }

        static public TimeSpan FromTicks(long value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromTicks(value));
        }

        public override int GetHashCode()
        { 
            return _timeSpan.GetHashCode(); 
        }

        //public TimeSpan Multiply(double factor)
        //{
        //    return new TimeSpan(_timeSpan.Multiply(factor));
        //}

        public TimeSpan Negate()
        {
            return new TimeSpan(_timeSpan.Negate());
        }

        static public TimeSpan Parse(string s)
        {
            return new TimeSpan(GSystem.TimeSpan.Parse(s));
        }

        static public TimeSpan Parse(string input, IFormatProvider formatProvider)
        {
            return new TimeSpan(GSystem.TimeSpan.Parse(input, DateTimeFormatInfo.Unwrap(formatProvider)));
        }

        static public TimeSpan ParseExact(string input, string format, IFormatProvider formatProvider)
        {
            return new TimeSpan(GSystem.TimeSpan.ParseExact(input,format, DateTimeFormatInfo.Unwrap(formatProvider)));
        }

        static public TimeSpan ParseExact(string input, string[] formats, IFormatProvider formatProvider)
        {
            return new TimeSpan(GSystem.TimeSpan.ParseExact(input, formats, DateTimeFormatInfo.Unwrap(formatProvider)));
        }

        static public TimeSpan ParseExact(string input, string format, IFormatProvider formatProvider, GSystem.Globalization.TimeSpanStyles styles)
        {
            return new TimeSpan(GSystem.TimeSpan.ParseExact(input, format, DateTimeFormatInfo.Unwrap(formatProvider), styles));
        }

        static public TimeSpan ParseExact(string input, string[] formats, IFormatProvider formatProvider, GSystem.Globalization.TimeSpanStyles styles)
        {
            return new TimeSpan(GSystem.TimeSpan.ParseExact(input, formats, DateTimeFormatInfo.Unwrap(formatProvider), styles));
        }

        public TimeSpan Subtract(TimeSpan ts)
        {
            return new TimeSpan(_timeSpan.Subtract(ts.WrappedTimeSpan));
        }

        public override string ToString()
        {
            return _timeSpan.ToString();
        }

        public string ToString2(string format)
        {
            return _timeSpan.ToString(format);
        }

        public string ToString3(string format, IFormatProvider formatProvider)
        {
            return _timeSpan.ToString(format, DateTimeFormatInfo.Unwrap(formatProvider));
        }

        public static bool TryParse(string s, out TimeSpan result)
        {
            bool tryParse = GSystem.TimeSpan.TryParse(s, out GSystem.TimeSpan outResult);
            result = new TimeSpan(outResult);
            return tryParse;
        }

        public static bool TryParse(string input, IFormatProvider formatProvider, out TimeSpan result)
        {
            bool tryParse = GSystem.TimeSpan.TryParse(input, DateTimeFormatInfo.Unwrap(formatProvider), out GSystem.TimeSpan outResult);
            result = new TimeSpan(outResult);
            return tryParse;
        }

        public static bool TryParseExact(string input, string format, IFormatProvider formatProvider, out TimeSpan result)
        {
            bool tryParseExact = GSystem.TimeSpan.TryParseExact(input, format, DateTimeFormatInfo.Unwrap(formatProvider), out GSystem.TimeSpan outResult);
            result = new TimeSpan(outResult);
            return tryParseExact;
        }

        public static  bool TryParseExact(string input, [In] string[] formats, IFormatProvider formatProvider, out TimeSpan result)
        {
            bool tryParseExact = GSystem.TimeSpan.TryParseExact(input, formats, DateTimeFormatInfo.Unwrap(formatProvider), out GSystem.TimeSpan outResult);
            result = new TimeSpan(outResult);
            return tryParseExact;
        }

        public static bool TryParseExact(string input, string format, IFormatProvider formatProvider, GSystem.Globalization.TimeSpanStyles styles, out TimeSpan result)
        {
            bool tryParseExact = GSystem.TimeSpan.TryParseExact(input, format, DateTimeFormatInfo.Unwrap(formatProvider), styles, out GSystem.TimeSpan outResult);
            result = new TimeSpan(outResult);
            return tryParseExact;
        }

        public static bool TryParseExact(string input, [In] string[] formats, IFormatProvider formatProvider, GSystem.Globalization.TimeSpanStyles styles, out TimeSpan result)
        {
            bool tryParseExact = GSystem.TimeSpan.TryParseExact(input, formats, DateTimeFormatInfo.Unwrap(formatProvider), styles, out GSystem.TimeSpan outResult);
            result = new TimeSpan(outResult);
            return tryParseExact;
        }

        // Operators

        public static TimeSpan Addition(TimeSpan t1, TimeSpan t2)
        {
            return new TimeSpan(t1.WrappedTimeSpan + t2.WrappedTimeSpan);
        }

        public static bool Equality(TimeSpan t1, TimeSpan t2)
        {
            return (t1.WrappedTimeSpan == t2.WrappedTimeSpan);
        }

        public static bool GreaterThan(TimeSpan t1, TimeSpan t2)
        {
            return (t1.WrappedTimeSpan > t2.WrappedTimeSpan);
        }

        public static bool GreaterThanOrEqual(TimeSpan t1, TimeSpan t2)
        {
            return (t1.WrappedTimeSpan >= t2.WrappedTimeSpan);
        }
        public static bool Inequality(TimeSpan t1, TimeSpan t2)
        {
            return (t1.WrappedTimeSpan != t2.WrappedTimeSpan);
        }

        public static bool LessThan(TimeSpan t1, TimeSpan t2)
        {
            return (t1.WrappedTimeSpan < t2.WrappedTimeSpan);
        }
        
        public static bool LessThanOrEqual(TimeSpan t1, TimeSpan t2)
        {
            return (t1.WrappedTimeSpan <= t2.WrappedTimeSpan);
        }

        public static TimeSpan Subtraction(TimeSpan t1, TimeSpan t2)
        {
            return new TimeSpan(t1.WrappedTimeSpan - t2.WrappedTimeSpan);
        }

        static public TimeSpan UnaryNegation(TimeSpan ts)
        {
            return new TimeSpan(-ts.WrappedTimeSpan);
        }

        static public TimeSpan UnaryPlus(TimeSpan ts)
        {
            return new TimeSpan(+ts.WrappedTimeSpan);
        }
    }
}
