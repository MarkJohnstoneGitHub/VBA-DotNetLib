// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.japanesecalendar?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("B5002641-5BDC-4DF0-B9F5-A056BDBBFE34")]
    [Description("Represents the Japanese calendar.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IJapaneseCalendarSingleton : ICalendarSingleton
    {
        [Description("Initializes a new instance of the JapaneseCalendar class.")]
        JapaneseCalendar Create();

        // Fields
        new int CurrentEra
        {
            [Description("Represents the current era of the current calendar. The value of this field is 0.")]
            get;
        }

        // Methods
        [Description("Determines whether the specified object instances are considered equal.")]
        new bool Equals(object objA, object objB);

        [Description("Returns a read-only version of the specified Calendar object.")]
        new Calendar ReadOnly(Calendar pCalendar);

    }
}
