// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.umalquracalendar?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("B4C1498A-8FAB-4608-B73D-542906AB533A")]
    [Description("Represents the Saudi Hijri (Um Al Qura) calendar.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IUmAlQuraCalendarSingleton : ICalendarSingleton
    {
        // Factory Methods
        [Description("Initializes a new instance of the UmAlQuraCalendar class.")]
        UmAlQuraCalendar Create();

        // Fields
        new int CurrentEra
        {
            [Description("Represents the current era of the current calendar. The value of this field is 0.")]
            get;
        }

        int UmAlQuraEra
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
