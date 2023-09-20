using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("7D2869A4-2585-4D71-9237-D00BCB5334C6")]
    [Description("Represents the Hebrew calendar.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IHebrewCalendarSingleton : ICalendarSingleton
    {
        // Factory Methods

        [Description("Initializes a new instance of the HebrewCalendar class.")]
        HebrewCalendar Create();

        // Fields

        new int CurrentEra
        {
            [Description("Represents the current era of the current calendar. The value of this field is 0.")]
            get;
        }

        int HebrewEra
        {
            [Description("Represents the current era. This field is constant.")]
            get;
        }

        // Methods

        [Description("Determines whether the specified object instances are considered equal.")]
        new bool Equals(object objA, object objB);

        [Description("Returns a read-only version of the specified Calendar object.")]
        new ICalendar ReadOnly(ICalendar pCalendar);
    }
}
