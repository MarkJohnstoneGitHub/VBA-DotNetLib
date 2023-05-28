using GSystem = global::System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System;

namespace DotNetLib.System
{

    // https://learn.microsoft.com/en-us/dotnet/api/system.timespan?view=netframework-4.8
    // https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/TimeSpan.cs

    [ComVisible(true)]
    [Guid("B73DFD69-6C69-4CFC-89F2-1C344228A9D4")]
    [ProgId("DotNetLib.System.TimeSpan")]
    [Description("Represents a time interval.")]
    [ClassInterface(ClassInterfaceType.None)]
    public class TimeSpan : ITimeSpan
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

        public TimeSpan(GSystem.TimeSpan value)
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

        public TimeSpan Create2(int days, int hours, int minutes, int seconds, int milliseconds)
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
            set { this.timeSpanObject = value; }  // set method
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
            return this.timeSpanObject.Equals(obj); 
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

        public override int GetHashCode()
        { 
            return this.timeSpanObject.GetHashCode(); 
        }

        public TimeSpan Negate()
        {
            return new TimeSpan(this.timeSpanObject.Negate());
        }

        public TimeSpan Parse(string s)
        {
            return new TimeSpan(GSystem.TimeSpan.Parse(s));
        }

        public TimeSpan Subtract(TimeSpan ts)
        {
            return new TimeSpan(this.timeSpanObject.Subtract(ts.timeSpanObject));
        }

        public string ToString(string format = null)
        {
            return this.timeSpanObject.ToString(format);
        }

        public bool TryParse(string s, out TimeSpan result)
        {
            bool pvtTryParse = GSystem.TimeSpan.TryParse(s, out GSystem.TimeSpan outResult);
            result = new TimeSpan(outResult);
            return pvtTryParse;
        }


    }
}
