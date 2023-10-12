using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    // https://learn.microsoft.com/en-us/dotnet/api/system.timespan?view=netframework-4.8

    [ComVisible(true)]
    [Guid("0844C467-32F4-489F-B286-9EAEE9F06CD9")]
    [Description("Represents a time interval.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITimeSpan
    {
        // Properties

        int Days
        {
            [Description("Gets the days component of the time interval represented by the current TimeSpan structure.")]
            get;
        }

        int Hours
        {
            [Description("Gets the hours component of the time interval represented by the current TimeSpan structure.")]
            get;
        }

        int Milliseconds
        {
            [Description("Gets the milliseconds component of the time interval represented by the current TimeSpan structure.")]
            get;
        }

        int Minutes
        {
            [Description("Gets the minutes component of the time interval represented by the current TimeSpan structure.")]
            get;
        }

        int Seconds
        {
            [Description("Gets the seconds component of the time interval represented by the current TimeSpan structure.")]
            get;
        }

        long Ticks
        {
            [Description("Gets the number of ticks that represent the value of the current TimeSpan structure.")]
            get;
        }

        double TotalDays
        {
            [Description("Gets the value of the current TimeSpan structure expressed in whole and fractional days.")]
            get;
        }

        double TotalHours
        {
            [Description("Gets the value of the current TimeSpan structure expressed in whole and fractional hours.")]
            get;
        }

        double TotalMilliseconds
        {
            [Description("Gets the value of the current TimeSpan structure expressed in whole and fractional milliseconds.")]
            get;
        }

        double TotalMinutes
        {
            [Description("Gets the value of the current TimeSpan structure expressed in whole and fractional minutes.")]
            get;
        }

        double TotalSeconds
        {
            [Description("Gets the value of the current TimeSpan structure expressed in whole and fractional seconds.")]
            get;
        }

        [Description("Returns a new TimeSpan object whose value is the sum of the specified TimeSpan object and this instance.")]
        TimeSpan Add(TimeSpan ts);

        [Description("Compares this instance to a specified TimeSpan object and returns an integer that indicates whether this instance is shorter than, equal to, or longer than the TimeSpan object.")]
        int CompareTo(TimeSpan value);

        [Description("Compares this instance to a specified object and returns an integer that indicates whether this instance is shorter than, equal to, or longer than the specified object.")]
        int CompareTo2(object value);

        [Description("Returns a new TimeSpan object whose value is the absolute value of the current TimeSpan object.")]
        TimeSpan Duration();

        [Description("Returns a value indicating whether this instance is equal to a specified TimeSpan object.")]
        bool Equals(TimeSpan obj);

        [Description("Returns a value indicating whether this instance is equal to a specified object.")]
        bool Equals2(object value);

        [Description("Returns a hash code for this instance.")]
        int GetHashCode();

        [Description("Returns a new TimeSpan object whose value is the negated value of this instance.")]
        TimeSpan Negate();

        [Description("Returns a new TimeSpan object whose value is the difference between the specified TimeSpan object and this instance.")]
        TimeSpan Subtract(TimeSpan ts);

        [Description("Converts the value of the current TimeSpan object to its equivalent string representation.")]
        string ToString();

        [Description("Converts the value of the current TimeSpan object to its equivalent string representation by using the specified format and culture-specific formatting information.")]
        string ToString2(string format, IFormatProvider formatProvider = null);

        //[Description("Converts the value of the current TimeSpan object to its equivalent string representation by using the specified format.")]
        //string ToString2(string format);

        //[Description("Converts the value of the current TimeSpan object to its equivalent string representation by using the specified format and culture-specific formatting information.")]
        //string ToString3(string format, IFormatProvider formatProvider);


        [Description("Gets the Type of the current instance.")]
        Type GetType();

    }
}
