// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.gregoriancalendar?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("30B05877-DCD1-4C7A-8F50-7F52469A7E3D")]
    [Description("Represents the Gregorian calendar.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IGregorianCalendarSingleton : ICalendarSingleton
    {
        // Factory Methods

        [Description("Initializes a new instance of the GregorianCalendar class using the default or specified GregorianCalendarTypes value.")]
        GregorianCalendar Create(GregorianCalendarTypes type = GregorianCalendarTypes.Localized);

        // Fields

        new int CurrentEra
        {
            [Description("Represents the current era of the current calendar. The value of this field is 0.")]
            get;
        }

        int ADEra
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
