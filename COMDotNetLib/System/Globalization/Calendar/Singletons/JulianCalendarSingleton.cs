using GGlobalization = global::System.Globalization;
using DotNetLib.Extensions;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("B2D41C01-218B-4458-A0F1-5AC43C8EDA2D")]
    [ProgId("DotNetLib.System.Globalization.JulianCalendarSingleton")]
    [Description("Represents the Julian calendar.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IJulianCalendarSingleton))]
    public class JulianCalendarSingleton : IJulianCalendarSingleton
    {

        public JulianCalendarSingleton() { }

        // Factory Methods
        public JulianCalendar Create()
        {
            return new JulianCalendar();
        }

        //Fields
        public int CurrentEra => GGlobalization.Calendar.CurrentEra;

        public int JulianEra => GGlobalization.JulianCalendar.JulianEra;

        // Methods
        public new bool Equals(object objA, object objB)
        {
            return GGlobalization.JulianCalendar.Equals(objA.Unwrap(), objB.Unwrap());
        }

        public Calendar ReadOnly(Calendar pCalendar)
        {
            return GGlobalization.Calendar.ReadOnly((GGlobalization.Calendar)pCalendar.Unwrap()).Wrap();
        }
    }
}
