// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.thaibuddhistcalendar?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("315208F8-8D4A-4A0B-977B-05B95AB9FA18")]
    [Description("")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IThaiBuddhistCalendarSingleton : ICalendarSingleton
    {
        // Factory Methods
        [Description("Initializes a new instance of the ThaiBuddhistCalendar class.")]
        ThaiBuddhistCalendar Create();

        // Fields
        new int CurrentEra
        {
            [Description("Represents the current era of the current calendar. The value of this field is 0.")]
            get;
        }

        int ThaiBuddhistEra
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
