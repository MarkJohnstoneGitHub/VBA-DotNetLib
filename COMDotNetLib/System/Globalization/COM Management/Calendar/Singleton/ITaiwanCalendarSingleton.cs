// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.taiwancalendar?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("50C938E8-8261-4713-8C08-4019799A5B6A")]
    [Description("Represents the Taiwan calendar.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITaiwanCalendarSingleton : ICalendarSingleton
    {
        // Factory Methods

        [Description("Initializes a new instance of the TaiwanCalendar class.")]
        TaiwanCalendar Create();

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
        new ICalendar ReadOnly(ICalendar calendar);

    }
}
