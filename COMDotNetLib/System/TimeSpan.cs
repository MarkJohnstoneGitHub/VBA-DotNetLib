using GSystem = global::System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System;
using DotNetLib.System.Globalization;

namespace DotNetLib.System
{

    // https://learn.microsoft.com/en-us/dotnet/api/system.timespan?view=netframework-4.8
    // https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/TimeSpan.cs

    [ComVisible(true)]
    [Guid("B73DFD69-6C69-4CFC-89F2-1C344228A9D4")]
    [ProgId("DotNetLib.System.TimeSpan")]
    [Description("Represents a time interval.")]
    [ClassInterface(ClassInterfaceType.None)]
    public class TimeSpan : IComTimeSpan
    {
        private GSystem.TimeSpan timeSpanObject;

        private static readonly TimeSpan tsMaxValue = new TimeSpan(GSystem.TimeSpan.MaxValue);
        private static readonly TimeSpan tsMinValue = new TimeSpan(GSystem.TimeSpan.MinValue);
        private static readonly TimeSpan tsZero = new TimeSpan(GSystem.TimeSpan.Zero);

        //Constructors
        public TimeSpan()
        {
            timeSpanObject = new GSystem.TimeSpan();
        }

        internal TimeSpan(GSystem.TimeSpan value)
        {
            this.timeSpanObject = value;
        }

        public TimeSpan(long ticks)
        {
            this.timeSpanObject = new GSystem.TimeSpan(ticks);
        }

        public TimeSpan(int hours, int minutes, int seconds)
        {
            this.timeSpanObject = new GSystem.TimeSpan(hours, minutes, seconds);
        }

        public TimeSpan(int days, int hours, int minutes, int seconds)
        {
            this.timeSpanObject = new GSystem.TimeSpan(days, hours, minutes, seconds);
        }

        public TimeSpan(int days, int hours, int minutes, int seconds, int milliseconds)
        {
            this.timeSpanObject = new GSystem.TimeSpan(days, hours, minutes, seconds, milliseconds);   
        }

        public TimeSpan CreateFromTicks(long ticks)
        {
            return new TimeSpan(ticks);
        }

        public TimeSpan Create(int hours, int minutes, int seconds)
        {
            return new TimeSpan(hours, minutes, seconds);
        }

        public TimeSpan Create2(int days, int hours, int minutes, int seconds)
        {
            return new TimeSpan(days, hours, minutes, seconds);
        }

        public TimeSpan Create3(int days, int hours, int minutes, int seconds, int milliseconds)
        {
            return new TimeSpan(days, hours, minutes, seconds, milliseconds);
        }

        // Fields

        public TimeSpan MaxValue => tsMaxValue;
        
        public TimeSpan MinValue => tsMinValue;

        public long TicksPerDay => GSystem.TimeSpan.TicksPerDay;
        
        public long TicksPerHour => GSystem.TimeSpan.TicksPerHour;

        public long TicksPerMillisecond => GSystem.TimeSpan.TicksPerMillisecond;

        public long TicksPerMinute => GSystem.TimeSpan.TicksPerMinute;

        public long TicksPerSecond => GSystem.TimeSpan.TicksPerSecond;

        public TimeSpan Zero => tsZero;


        //Properties
        internal GSystem.TimeSpan TimeSpanObject
        {
            get { return this.timeSpanObject; }
            set { this.timeSpanObject = value; } 
        }

        public int Days => this.timeSpanObject.Days;

        public int Hours => this.timeSpanObject.Hours;

        public int Milliseconds => this.timeSpanObject.Milliseconds;

        public int Minutes => this.timeSpanObject.Minutes;

        public int Seconds => this.timeSpanObject.Seconds;

        public long Ticks => this.timeSpanObject.Ticks;

        public double TotalDays => this.timeSpanObject.TotalDays;

        public double TotalHours => this.timeSpanObject.TotalHours;

        public double TotalMilliseconds => this.timeSpanObject.TotalMilliseconds;
        
        public double TotalMinutes => this.timeSpanObject.TotalMinutes;
        
        public double TotalSeconds => this.timeSpanObject.TotalSeconds;


        //Methods
        public TimeSpan Add(TimeSpan ts)
        {
            return new TimeSpan(this.timeSpanObject.Add(ts.timeSpanObject));
        }

        public int Compare(TimeSpan t1, TimeSpan t2)
        {
            return GSystem.TimeSpan.Compare(t1.timeSpanObject, t2.timeSpanObject);  
        }

        public int CompareTo(TimeSpan value)
        {
            return this.timeSpanObject.CompareTo(value.timeSpanObject);
        }

        // TODO : Check implementation
        public int CompareTo2(object value)
        {
            if (value == null) return 1;
            return this.timeSpanObject.CompareTo((TimeSpan)value);

            //if (value == null) return 1;
            //if (!(value is TimeSpan))
            //    throw new ArgumentException(SR.Arg_MustBeTimeSpan);
            //long t = ((TimeSpan)value)._ticks;
            //if (_ticks > t) return 1;
            //if (_ticks < t) return -1;
            //return 0;
        }

        public TimeSpan Duration()
        {
            return new TimeSpan(this.timeSpanObject.Duration());
        }

        public bool Equals(TimeSpan obj)
        { 
            return this.timeSpanObject.Equals(obj.TimeSpanObject); 
        }

        // TODO : Check Implementation
        public bool Equals2(object value)
        {
            return value is TimeSpan ts && this.timeSpanObject == ts.timeSpanObject;

            //if (value is TimeSpan)
            //{
            //    return this.timeSpanObject.Equals((TimeSpan)value);
            //}
            //return false;
        }

        public bool Equals3(TimeSpan t1, TimeSpan t2)
        { 
            return GSystem.TimeSpan.Equals(t1.timeSpanObject,t2.timeSpanObject); 
        }

        public TimeSpan FromDays(double value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromDays(value));
        }

        public TimeSpan FromHours(double value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromHours(value));
        }

        public TimeSpan FromMilliseconds(double value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromMilliseconds(value));
        }

        public TimeSpan FromMinutes(double value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromMinutes(value));
        }

        public TimeSpan FromSeconds(double value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromSeconds(value));
        }

        public TimeSpan FromTicks(long value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromTicks(value));
        }

        public override int GetHashCode()
        { 
            return this.timeSpanObject.GetHashCode(); 
        }

        //public TimeSpan Multiply(double factor)
        //{
        //    return new TimeSpan(this.timeSpanObject.Multiply(factor));
        //}
        public TimeSpan Negate()
        {
            return new TimeSpan(this.timeSpanObject.Negate());
        }

        public TimeSpan Parse(string s)
        {
            return new TimeSpan(GSystem.TimeSpan.Parse(s));
        }

        public TimeSpan Parse2(string input, GSystem.IFormatProvider formatProvider)
        {
            return new TimeSpan(GSystem.TimeSpan.Parse(input,formatProvider));
        }

        public TimeSpan ParseExact(string input, string format, GSystem.IFormatProvider formatProvider)
        {
            return new TimeSpan(GSystem.TimeSpan.ParseExact(input,format, formatProvider));
        }

        public TimeSpan ParseExact2(string input, [In] ref string[] formats, GSystem.IFormatProvider formatProvider)
        {
            return new TimeSpan(GSystem.TimeSpan.ParseExact(input, formats, formatProvider));
        }

        public TimeSpan ParseExact3(string input, string format, GSystem.IFormatProvider formatProvider, TimeSpanStyles styles)
        {
            return new TimeSpan(GSystem.TimeSpan.ParseExact(input, format, formatProvider, (GSystem.Globalization.TimeSpanStyles)styles));
        }

        public TimeSpan ParseExact4(string input, [In] ref string[] formats, GSystem.IFormatProvider formatProvider, TimeSpanStyles styles)
        {
            return new TimeSpan(GSystem.TimeSpan.ParseExact(input, formats, formatProvider, (GSystem.Globalization.TimeSpanStyles)styles));
        }

        public TimeSpan Subtract(TimeSpan ts)
        {
            return new TimeSpan(this.timeSpanObject.Subtract(ts.timeSpanObject));
        }


        public override string ToString()
        {
            return this.timeSpanObject.ToString();
        }

        public string ToString2(string format)
        {
            return this.timeSpanObject.ToString(format);
        }

        public string ToString3(string format, GSystem.IFormatProvider formatProvider)
        {
            return this.timeSpanObject.ToString(format, formatProvider);
        }

        public bool TryParse(string s, out TimeSpan result)
        {
            bool tryParse = GSystem.TimeSpan.TryParse(s, out GSystem.TimeSpan outResult);
            result = new TimeSpan(outResult);
            return tryParse;
        }

        public bool TryParse2(string input, GSystem.IFormatProvider formatProvider, out TimeSpan result)
        {
            bool tryParse = GSystem.TimeSpan.TryParse(input, formatProvider, out GSystem.TimeSpan outResult);
            result = new TimeSpan(outResult);
            return tryParse;
        }

        public bool TryParseExact(string input, string format, IFormatProvider formatProvider, out TimeSpan result)
        {
            bool tryParseExact = GSystem.TimeSpan.TryParseExact(input, format, formatProvider, out GSystem.TimeSpan outResult);
            result = new TimeSpan(outResult);
            return tryParseExact;
        }

        public  bool TryParseExact2(string input, [In] ref string[] formats, IFormatProvider formatProvider, out TimeSpan result)
        {
            bool tryParseExact = GSystem.TimeSpan.TryParseExact(input, formats, formatProvider, out GSystem.TimeSpan outResult);
            result = new TimeSpan(outResult);
            return tryParseExact;
        }

        public bool TryParseExact3(string input, string format, IFormatProvider formatProvider, TimeSpanStyles styles, out TimeSpan result)
        {
            bool tryParseExact = GSystem.TimeSpan.TryParseExact(input, format, formatProvider, (GSystem.Globalization.TimeSpanStyles)styles, out GSystem.TimeSpan outResult);
            result = new TimeSpan(outResult);
            return tryParseExact;
        }

        public bool TryParseExact4(string input, [In] ref string[] formats, IFormatProvider formatProvider, TimeSpanStyles styles, out TimeSpan result)
        {
            bool tryParseExact = GSystem.TimeSpan.TryParseExact(input, formats, formatProvider, (GSystem.Globalization.TimeSpanStyles)styles, out GSystem.TimeSpan outResult);
            result = new TimeSpan(outResult);
            return tryParseExact;
        }

        // Operators

        public TimeSpan Addition(TimeSpan t1, TimeSpan t2)
        {
            return new TimeSpan(t1.timeSpanObject + t2.TimeSpanObject);
        }

        public bool Equality(TimeSpan t1, TimeSpan t2)
        {
            return (t1.timeSpanObject == t2.timeSpanObject);
        }

        public bool GreaterThan(TimeSpan t1, TimeSpan t2)
        {
            return (t1.timeSpanObject > t2.timeSpanObject);
        }

        public bool GreaterThanOrEqual(TimeSpan t1, TimeSpan t2)
        {
            return (t1.timeSpanObject >= t2.timeSpanObject);
        }
        public bool Inequality(TimeSpan t1, TimeSpan t2)
        {
            return (t1.timeSpanObject != t2.timeSpanObject);
        }

        public bool LessThan(TimeSpan t1, TimeSpan t2)
        {
            return (t1.timeSpanObject < t2.timeSpanObject);
        }
        public bool LessThanOrEqual(TimeSpan t1, TimeSpan t2)
        {
            return (t1.timeSpanObject <= t2.timeSpanObject);
        }

        public TimeSpan Subtraction(TimeSpan t1, TimeSpan t2)
        {
            return new TimeSpan(t1.timeSpanObject - t2.timeSpanObject);
        }

        public TimeSpan UnaryNegation(TimeSpan t)
        {
            return new TimeSpan(-t.timeSpanObject);
        }

        public TimeSpan UnaryPlus(TimeSpan t)
        {
            return new TimeSpan(+ t.timeSpanObject);
        }


    }
}
