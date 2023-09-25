using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("A17B8B50-BE34-4D2B-8996-C37BFD308305")]
    [Description("Calendar static fields.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICalendarSingleton
    {
        int CurrentEra
        {
            [Description("Represents the current era of the current calendar. The value of this field is 0.")]
            get;
        }

        [Description("Determines whether the specified object instances are considered equal.")]
        bool Equals(object objA, object objB);

        [Description("Returns a read-only version of the specified Calendar object.")]
        Calendar ReadOnly(Calendar pCalendar);



    }
}
