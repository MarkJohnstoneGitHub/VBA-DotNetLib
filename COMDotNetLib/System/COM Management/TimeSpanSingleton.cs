// https://learn.microsoft.com/en-us/dotnet/api/system.timespan?view=netframework-4.8

using GSystem = global::System;
using System;
using System.Runtime.InteropServices;
using DotNetLib.System.Globalization;
using System.ComponentModel;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("2B0795AF-E311-4D8E-839E-304CC7618F58")]
    [ProgId("DotNetLib.System.TimeSpanSingleton")]
    [Description("TimeSpan factory methods and static members.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITimeSpanSingleton))]
    public class TimeSpanSingleton : ITimeSpanSingleton
    {
        public TimeSpanSingleton() { }

        // Factory Methods
        public TimeSpan CreateFromTicks(long ticks)
        {
            return new TimeSpan(ticks);
        }

        public TimeSpan Create(int pHours, int pMinutes, int pSeconds)
        {
            return new TimeSpan(pHours, pMinutes, pSeconds);
        }

        public TimeSpan Create2(int pDays, int pHours, int pMinutes, int pSeconds)
        {
            return new TimeSpan(pDays, pHours, pMinutes, pSeconds);
        }

        public TimeSpan Create3(int pDays, int pHours, int pMinutes, int pSeconds, int pMilliseconds)
        {
            return new TimeSpan(pDays, pHours, pMinutes, pSeconds, pMilliseconds);
        }

        // Fields

        public TimeSpan MaxValue => TimeSpan.MaxValue;

        public TimeSpan MinValue => TimeSpan.MinValue;

        public long TicksPerDay => TimeSpan.TicksPerDay;

        public long TicksPerHour => TimeSpan.TicksPerHour;

        public long TicksPerMillisecond => TimeSpan.TicksPerMillisecond;

        public long TicksPerMinute => TimeSpan.TicksPerMinute;

        public long TicksPerSecond => TimeSpan.TicksPerSecond;

        public TimeSpan Zero => TimeSpan.Zero;

        // Methods
        public int Compare(TimeSpan ts1, TimeSpan ts2)
        {
            return TimeSpan.Compare(ts1, ts2);
        }

        public bool Equals(TimeSpan t1, TimeSpan t2)
        {
            return TimeSpan.Equals(t1, t2);
        }

        public TimeSpan FromDays(double value)
        {
            return TimeSpan.FromDays(value);
        }

        public TimeSpan FromHours(double value)
        {
            return TimeSpan.FromHours(value);
        }

        public TimeSpan FromMilliseconds(double value)
        {
            return TimeSpan.FromMilliseconds(value);
        }

        public TimeSpan FromMinutes(double value)
        {
            return TimeSpan.FromMinutes(value);
        }

        public TimeSpan FromSeconds(double value)
        {
            return TimeSpan.FromSeconds(value);
        }

        public TimeSpan FromTicks(long value)
        {
            return TimeSpan.FromTicks(value);
        }

        public TimeSpan Parse(string s)
        {
            return TimeSpan.Parse(s);
        }

        public TimeSpan Parse2(string input, GSystem.IFormatProvider formatProvider)
        {
            return TimeSpan.Parse(input, formatProvider);
        }

        public TimeSpan ParseExact(string input, string format, GSystem.IFormatProvider formatProvider)
        {
            return TimeSpan.ParseExact(input, format, formatProvider);
        }

        public TimeSpan ParseExact2(string input, [In] ref string[] formats, GSystem.IFormatProvider formatProvider)
        {
            return TimeSpan.ParseExact(input, formats, formatProvider);
        }

        public TimeSpan ParseExact3(string input, string format, GSystem.IFormatProvider formatProvider, TimeSpanStyles styles)
        {
            return TimeSpan.ParseExact(input, format, formatProvider, (GSystem.Globalization.TimeSpanStyles)styles);
        }

        public TimeSpan ParseExact4(string input, [In] ref string[] formats, GSystem.IFormatProvider formatProvider, TimeSpanStyles styles)
        {
            return TimeSpan.ParseExact(input, formats, formatProvider, (GSystem.Globalization.TimeSpanStyles)styles);
        }

        public bool TryParse(string s, out TimeSpan result)
        {
            return TimeSpan.TryParse(s, out result);
        }

        public bool TryParse2(string input, GSystem.IFormatProvider formatProvider, out TimeSpan result)
        {
            return TimeSpan.TryParse(input, formatProvider,out result);
        }

        public bool TryParseExact(string input, string format, IFormatProvider formatProvider, out TimeSpan result)
        {
            return TimeSpan.TryParseExact(input, format,formatProvider, out result);
        }

        public bool TryParseExact2(string input, [In] ref string[] formats, IFormatProvider formatProvider, out TimeSpan result)
        {
            return TimeSpan.TryParseExact(input, formats, formatProvider, out result);
        }

        public bool TryParseExact3(string input, string format, IFormatProvider formatProvider, TimeSpanStyles styles, out TimeSpan result)
        {
            return TimeSpan.TryParseExact(input, format, formatProvider, (GSystem.Globalization.TimeSpanStyles)styles, out result);
        }

        public bool TryParseExact4(string input, [In] ref string[] formats, IFormatProvider formatProvider, TimeSpanStyles styles, out TimeSpan result)
        {
            return TimeSpan.TryParseExact(input, formats, formatProvider, (GSystem.Globalization.TimeSpanStyles)styles, out result);
        }

        // Operators

        public TimeSpan Addition(TimeSpan t1, TimeSpan t2)
        {
            return TimeSpan.Addition(t1, t2);
        }

        public bool Equality(TimeSpan t1, TimeSpan t2)
        {
            return TimeSpan.Equality(t1, t2);
        }

        public bool GreaterThan(TimeSpan t1, TimeSpan t2)
        {
            return TimeSpan.GreaterThan(t1,t2);
        }

        public bool GreaterThanOrEqual(TimeSpan t1, TimeSpan t2)
        {
            return TimeSpan.GreaterThanOrEqual(t1, t2);
        }
        public bool Inequality(TimeSpan t1, TimeSpan t2)
        {
            return TimeSpan.Inequality(t1,t2);
        }

        public bool LessThan(TimeSpan t1, TimeSpan t2)
        {
            return TimeSpan.LessThan(t1, t2);
        }

        public bool LessThanOrEqual(TimeSpan t1, TimeSpan t2)
        {
            return TimeSpan.LessThanOrEqual(t1, t2);
        }

        public TimeSpan Subtraction(TimeSpan t1, TimeSpan t2)
        {
            return TimeSpan.Subtraction(t1, t2);
        }

        public TimeSpan UnaryNegation(TimeSpan ts)
        {
            return TimeSpan.UnaryNegation(ts);
        }

        public TimeSpan UnaryPlus(TimeSpan ts)
        {
            return TimeSpan.UnaryPlus(ts);
        }
    }
}
