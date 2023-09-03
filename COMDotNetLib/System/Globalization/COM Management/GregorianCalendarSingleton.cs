using GGlobalization = global::System.Globalization;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{

    [ComVisible(false)]
    [Guid("69AFAFEA-385A-48CE-91C0-790459722718")]
    [ProgId("DotNetLib.System.Globalization.GregorianCalendarSingleton")]
    [Description("Represents the Gregorian calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IGregorianCalendarSingleton))]
    public class GregorianCalendarSingleton : IGregorianCalendarSingleton
    {
        public GregorianCalendarSingleton() { }


        public GregorianCalendar Create()
        {
            return new GregorianCalendar();
        }

        public GregorianCalendar Create2(GregorianCalendarTypes type)
        {
            return new GregorianCalendar(type);
        }
        public GregorianCalendar CreateFromGregorianCalendar(GGlobalization.GregorianCalendar gregorianCalendar)
        {
            return new GregorianCalendar(gregorianCalendar);
        }




    }
}
