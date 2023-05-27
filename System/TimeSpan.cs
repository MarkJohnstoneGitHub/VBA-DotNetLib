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
        private GSystem.TimeSpan objTimeSpan;

        //Constructors
        public TimeSpan()
        {
            objTimeSpan = new GSystem.TimeSpan();
        }

        public TimeSpan(GSystem.TimeSpan value)
        {
            this.objTimeSpan = value;
        }

        public TimeSpan(long ticks)
        {
            this.objTimeSpan = new GSystem.TimeSpan(ticks);
        }

        public TimeSpan(int hours, int minutes, int seconds)
        {
            this.objTimeSpan = new GSystem.TimeSpan(hours, minutes, seconds);
        }

        public TimeSpan(int days, int hours, int minutes, int seconds, int milliseconds)
        {
            this.objTimeSpan = new GSystem.TimeSpan(days, hours, minutes, seconds, milliseconds);   
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

        internal GSystem.TimeSpan timeSpan
        {
            get { return objTimeSpan; }
            set { objTimeSpan = value; }  // set method
        }


        //Properties
        public int Days => this.objTimeSpan.Days;
        public int Hours => this.objTimeSpan.Hours;

        public int Milliseconds => this.objTimeSpan.Milliseconds;
        public int Minutes => this.objTimeSpan.Minutes;
        public int Seconds => this.objTimeSpan.Seconds;
        public long Ticks => this.objTimeSpan.Ticks;
        public double TotalDays => this.objTimeSpan.TotalDays;
        public double TotalHours => this.objTimeSpan.TotalHours;
        public double TotalMinutes => this.objTimeSpan.TotalMinutes;
        public double TotalSeconds => this.objTimeSpan.TotalSeconds;
        public double TotalMilliseconds => this.objTimeSpan.TotalMilliseconds;

        //Methods
        public TimeSpan Add(TimeSpan ts)
        {
            return new TimeSpan(this.objTimeSpan.Add(ts.objTimeSpan));
        }

        public int Compare(TimeSpan t1, TimeSpan t2)
        {
            return GSystem.TimeSpan.Compare(t1.objTimeSpan, t2.objTimeSpan);  
        }

        public int CompareTo(TimeSpan value)
        {
            return this.objTimeSpan.CompareTo(value.objTimeSpan);
        }

        public TimeSpan Duration()
        {
            return new TimeSpan(this.objTimeSpan.Duration());
        }

        public bool Equals(TimeSpan obj)
        { 
            return this.objTimeSpan.Equals(obj); 
        }

        public bool Equals2(object value)
        {
            return this.objTimeSpan.Equals(value);
        }

        public bool Equals3(TimeSpan t1, TimeSpan t2)
        { 
            return GSystem.TimeSpan.Equals(t1.objTimeSpan,t2.objTimeSpan); 
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
            return this.objTimeSpan.GetHashCode(); 
        }

        public TimeSpan Negate()
        {
            return new TimeSpan(this.objTimeSpan.Negate());
        }

        public TimeSpan Parse(string s)
        {
            return new TimeSpan(GSystem.TimeSpan.Parse(s));
        }

        public TimeSpan Subtract(TimeSpan ts)
        {
            return new TimeSpan(this.objTimeSpan.Subtract(ts.objTimeSpan));
        }

        public string ToString(string format = null)
        {
            return this.objTimeSpan.ToString(format);
        }

        public bool TryParse(string s, out TimeSpan result)
        {
            bool pvtTryParse = GSystem.TimeSpan.TryParse(s, out GSystem.TimeSpan outResult);
            result = new TimeSpan(outResult);
            return pvtTryParse;
        }
    }
}
