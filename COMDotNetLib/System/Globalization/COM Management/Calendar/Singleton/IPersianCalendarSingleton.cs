// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.persiancalendar?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("ECF77460-D12F-4AE5-A183-A7CF2CAC4F09")]
    [Description("Represents the Persian calendar.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IPersianCalendarSingleton : ICalendarSingleton
    {
        [Description("Initializes a new instance of the PersianCalendar class.")]
        PersianCalendar Create();

        // Fields
        new int CurrentEra
        {
            [Description("Represents the current era of the current calendar. The value of this field is 0.")]
            get;
        }

        int PersianEra
        {
            [Description("Represents the current era. This field is constant.")]
            get;
        }


        // Methods
        [Description("Determines whether the specified object instances are considered equal.")]
        new bool Equals(object objA, object objB);

        [Description("Returns a read-only version of the specified Calendar object.")]
        new ICalendar ReadOnly(ICalendar calendar);

    }
}
