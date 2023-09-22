using GGlobalization = global::System.Globalization;
using DotNetLib.Extensions;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("69AFAFEA-385A-48CE-91C0-790459722718")]
    [ProgId("DotNetLib.System.Globalization.GregorianCalendarSingleton")]
    [Description("Represents the Gregorian calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IGregorianCalendarSingleton))]
    public class GregorianCalendarSingleton : IGregorianCalendarSingleton
    {
        public GregorianCalendarSingleton() { }


        // Factory Methods

        public GregorianCalendar Create(GregorianCalendarTypes type = GregorianCalendarTypes.Localized)
        {
            return new GregorianCalendar(type);
        }

        // Fields

        public int CurrentEra => GGlobalization.Calendar.CurrentEra;

        public int ADEra => GGlobalization.GregorianCalendar.ADEra;


        // Methods

        public new bool Equals(object objA, object objB)
        {
            return GGlobalization.GregorianCalendar.Equals(objA.Unwrap(), objB.Unwrap());
        }

        public Calendar ReadOnly(Calendar pCalendar)
        {
            return GGlobalization.Calendar.ReadOnly((GGlobalization.Calendar)pCalendar.Unwrap()).Wrap();
        }
    }
}
