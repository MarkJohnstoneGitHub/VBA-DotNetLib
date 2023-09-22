// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.hijricalendar?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("F6C7F9B9-027D-482A-A2AD-98C46301CA82")]
    [Description("Represents the Hijri calendar.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IHijriCalendarSingleton : ICalendarSingleton
    {
        // Factory Methods

        [Description("Initializes a new instance of the HijriCalendar class.")]
        HijriCalendar Create();

        // Fields

        new int CurrentEra
        {
            [Description("Represents the current era of the current calendar. The value of this field is 0.")]
            get;
        }

        int HijriEra
        {
            [Description("Represents the current era. This field is constant.")]
            get;
        }

        // Methods

        [Description("Determines whether the specified object instances are considered equal.")]
        new bool Equals(object objA, object objB);

        [Description("Returns a read-only version of the specified Calendar object.")]
        new Calendar ReadOnly(Calendar pCalendar);
    }
}
