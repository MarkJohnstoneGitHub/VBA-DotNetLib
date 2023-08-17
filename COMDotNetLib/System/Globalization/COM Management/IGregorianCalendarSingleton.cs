using GGlobalization = global::System.Globalization;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(false)]
    [Guid("30B05877-DCD1-4C7A-8F50-7F52469A7E3D")]
    [Description("Represents the Gregorian calendar.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IGregorianCalendarSingleton
    {
        GregorianCalendar Create();

        GregorianCalendar Create2(GregorianCalendarTypes type);

        GregorianCalendar CreateFromGregorianCalendar(GGlobalization.GregorianCalendar gregorianCalendar);

    }
}
