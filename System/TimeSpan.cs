using GSystem = global::System;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("B73DFD69-6C69-4CFC-89F2-1C344228A9D4")]
    [ProgId("DotNetLib.System.TimeSpan")]
    [ClassInterface(ClassInterfaceType.None)]
    public class TimeSpan : ITimeSpan
    {
        private GSystem.TimeSpan timeSpanObject;

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

        public readonly TimeSpan MaxValue;
        // public static readonly TimeSpan MinValue;
        // public const long TicksPerDay = 864000000000;
        // public const long TicksPerHour = 36000000000;
        // public const long TicksPerMillisecond = 10000;
        // public const long TicksPerMinute = 600000000;
        // public const long TicksPerSecond = 10000000;
        // public static readonly TimeSpan Zero;

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

        public TimeSpan Duration()
        {
            return new TimeSpan(this.timeSpanObject.Duration());
        }

        public bool Equals(TimeSpan obj)
        { 
            return this.timeSpanObject.Equals(obj); 
        }

        public bool Equals2(object value)
        {
            return this.timeSpanObject.Equals(value);
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
