using System;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{

    [ComVisible(true)]
    [Guid("C2F50F87-A7D4-4C32-A01B-EC750CE1B955")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITimeSpan
    {
        //Constructors
        TimeSpan CreateFromTicks(long ticks);
        TimeSpan Create(int hours, int minutes, int seconds);
        TimeSpan Create2(int days, int hours, int minutes, int seconds, int milliseconds);

        int Days { get; }
        int Hours { get; }
        int Milliseconds { get; }
        int Minutes { get; }
        int Seconds { get; }
        long Ticks { get; }
        double TotalDays { get; }
        double TotalHours { get; }
        double TotalMilliseconds { get; }
        double TotalMinutes { get; }
        double TotalSeconds { get; }

        TimeSpan Add(TimeSpan ts);
        int Compare(TimeSpan t1, TimeSpan t2);
        int CompareTo(TimeSpan value);
        TimeSpan Duration();
        bool Equals(TimeSpan obj);
        bool Equals2(object value);
        bool Equals3(TimeSpan t1, TimeSpan t2);
        TimeSpan FromDays(double value);
        TimeSpan FromHours(double value);
        TimeSpan FromMilliseconds(double value);
        TimeSpan FromMinutes(double value);
        TimeSpan FromSeconds(double value);
        int GetHashCode();
        TimeSpan Negate();
        TimeSpan Parse(string s);
        TimeSpan Subtract(TimeSpan ts);
        string ToString(string format = null);
        bool TryParse(string s, out TimeSpan result);
        

    }
}
