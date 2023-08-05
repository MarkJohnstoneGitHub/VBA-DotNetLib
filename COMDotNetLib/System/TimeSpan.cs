// https://learn.microsoft.com/en-us/dotnet/api/system.timespan?view=netframework-4.8
// https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/TimeSpan.cs

using GSystem = global::System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System;
using DotNetLib.System.Globalization;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("B73DFD69-6C69-4CFC-89F2-1C344228A9D4")]
    [ProgId("DotNetLib.System.TimeSpan")]
    [Description("Represents a time interval.")]
    [ClassInterface(ClassInterfaceType.None)]
    public class TimeSpan : ITimeSpan, ITimeSpanSingleton
    {
        private GSystem.TimeSpan timeSpanObject;

        // Static Fields
        private static readonly ITimeSpan tsMaxValue = new TimeSpan(GSystem.TimeSpan.MaxValue);
        private static readonly ITimeSpan tsMinValue = new TimeSpan(GSystem.TimeSpan.MinValue);
        private static readonly ITimeSpan tsZero = new TimeSpan(GSystem.TimeSpan.Zero);

        //Constructors
        public TimeSpan()
        {
            timeSpanObject = new GSystem.TimeSpan();
        }

        internal TimeSpan(GSystem.TimeSpan value)
        {
            this.timeSpanObject = value;
        }

        internal TimeSpan(long ticks)
        {
            this.timeSpanObject = new GSystem.TimeSpan(ticks);
        }

        internal TimeSpan(int hours, int minutes, int seconds)
        {
            this.timeSpanObject = new GSystem.TimeSpan(hours, minutes, seconds);
        }

        internal TimeSpan(int days, int hours, int minutes, int seconds)
        {
            this.timeSpanObject = new GSystem.TimeSpan(days, hours, minutes, seconds);
        }

        internal TimeSpan(int days, int hours, int minutes, int seconds, int milliseconds)
        {
            this.timeSpanObject = new GSystem.TimeSpan(days, hours, minutes, seconds, milliseconds);   
        }

        public ITimeSpan CreateFromTicks(long ticks)
        {
            return new TimeSpan(ticks);
        }

        public ITimeSpan Create(int pHours, int pMinutes, int pSeconds)
        {
            return new TimeSpan(pHours, pMinutes, pSeconds);
        }

        public ITimeSpan Create2(int pDays, int pHours, int pMinutes, int pSeconds)
        {
            return new TimeSpan(pDays, pHours, pMinutes, pSeconds);
        }

        public ITimeSpan Create3(int pDays, int pHours, int pMinutes, int pSeconds, int pMilliseconds)
        {
            return new TimeSpan(pDays, pHours, pMinutes, pSeconds, pMilliseconds);
        }

        // Fields

        public ITimeSpan MaxValue => tsMaxValue;
        
        public ITimeSpan MinValue => tsMinValue;

        public long TicksPerDay => GSystem.TimeSpan.TicksPerDay;
        
        public long TicksPerHour => GSystem.TimeSpan.TicksPerHour;

        public long TicksPerMillisecond => GSystem.TimeSpan.TicksPerMillisecond;

        public long TicksPerMinute => GSystem.TimeSpan.TicksPerMinute;

        public long TicksPerSecond => GSystem.TimeSpan.TicksPerSecond;

        public ITimeSpan Zero => tsZero;


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
        public ITimeSpan Add(TimeSpan ts)
        {
            return new TimeSpan(this.timeSpanObject.Add(ts.TimeSpanObject));
        }

        public int Compare(TimeSpan ts1, TimeSpan ts2)
        {
            return GSystem.TimeSpan.Compare(ts1.TimeSpanObject,ts2.TimeSpanObject);  
        }

        public int CompareTo(TimeSpan value)
        {
            return this.timeSpanObject.CompareTo(value.TimeSpanObject);
        }

        public int CompareTo2(object value)
        {
            const string Arg_MustBeTimeSpan = "Object must be of type TimeSpan.";

            if (value == null) return 1;
            if (!(value is TimeSpan ts))
            {
                throw new ArgumentException(Arg_MustBeTimeSpan);
            }
            return this.CompareTo(ts);
        }

        public ITimeSpan Duration()
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
            return value is TimeSpan ts && this.timeSpanObject == ts.TimeSpanObject;
        }

        public bool Equals3(TimeSpan t1, TimeSpan t2)
        { 
            return GSystem.TimeSpan.Equals(t1.TimeSpanObject, t2.TimeSpanObject); 
        }

        public ITimeSpan FromDays(double value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromDays(value));
        }

        public ITimeSpan FromHours(double value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromHours(value));
        }

        public ITimeSpan FromMilliseconds(double value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromMilliseconds(value));
        }

        public ITimeSpan FromMinutes(double value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromMinutes(value));
        }

        public ITimeSpan FromSeconds(double value)
        {
            return new TimeSpan(GSystem.TimeSpan.FromSeconds(value));
        }

        public ITimeSpan FromTicks(long value)
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
        public ITimeSpan Negate()
        {
            return new TimeSpan(this.timeSpanObject.Negate());
        }

        public ITimeSpan Parse(string s)
        {
            return new TimeSpan(GSystem.TimeSpan.Parse(s));
        }

        public ITimeSpan Parse2(string input, GSystem.IFormatProvider formatProvider)
        {
            return new TimeSpan(GSystem.TimeSpan.Parse(input,formatProvider));
        }

        public ITimeSpan ParseExact(string input, string format, GSystem.IFormatProvider formatProvider)
        {
            return new TimeSpan(GSystem.TimeSpan.ParseExact(input,format, formatProvider));
        }

        public ITimeSpan ParseExact2(string input, [In] ref string[] formats, GSystem.IFormatProvider formatProvider)
        {
            return new TimeSpan(GSystem.TimeSpan.ParseExact(input, formats, formatProvider));
        }

        public ITimeSpan ParseExact3(string input, string format, GSystem.IFormatProvider formatProvider, TimeSpanStyles styles)
        {
            return new TimeSpan(GSystem.TimeSpan.ParseExact(input, format, formatProvider, (GSystem.Globalization.TimeSpanStyles)styles));
        }

        public ITimeSpan ParseExact4(string input, [In] ref string[] formats, GSystem.IFormatProvider formatProvider, TimeSpanStyles styles)
        {
            return new TimeSpan(GSystem.TimeSpan.ParseExact(input, formats, formatProvider, (GSystem.Globalization.TimeSpanStyles)styles));
        }

        public ITimeSpan Subtract(TimeSpan ts)
        {
            return new TimeSpan(this.timeSpanObject.Subtract(ts.TimeSpanObject));
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
            return new TimeSpan(t1.TimeSpanObject + t2.TimeSpanObject);
        }

        public bool Equality(TimeSpan t1, TimeSpan t2)
        {
            return (t1.TimeSpanObject == t2.TimeSpanObject);
        }

        public bool GreaterThan(TimeSpan t1, TimeSpan t2)
        {
            return (t1.TimeSpanObject > t2.TimeSpanObject);
        }

        public bool GreaterThanOrEqual(TimeSpan t1, TimeSpan t2)
        {
            return (t1.TimeSpanObject >= t2.TimeSpanObject);
        }
        public bool Inequality(TimeSpan t1, TimeSpan t2)
        {
            return (t1.TimeSpanObject != t2.TimeSpanObject);
        }

        public bool LessThan(TimeSpan t1, TimeSpan t2)
        {
            return (t1.TimeSpanObject < t2.TimeSpanObject);
        }
        public bool LessThanOrEqual(TimeSpan t1, TimeSpan t2)
        {
            return (t1.TimeSpanObject <= t2.TimeSpanObject);
        }

        public ITimeSpan Subtraction(TimeSpan t1, TimeSpan t2)
        {
            return new TimeSpan(t1.TimeSpanObject - t2.TimeSpanObject);
        }

        public ITimeSpan UnaryNegation(TimeSpan ts)
        {
            return new TimeSpan(-ts.TimeSpanObject);
        }

        public ITimeSpan UnaryPlus(TimeSpan ts)
        {
            return new TimeSpan(+ts.TimeSpanObject);
        }


    }
}
