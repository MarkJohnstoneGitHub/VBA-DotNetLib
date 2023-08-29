using System.ComponentModel;
using System;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("15147227-474D-4DF2-9595-4410072F8C02")]
    [Description("DateTimeFormatInfo factory methods and static members that provides culture-specific information about the format of date and time values.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDateTimeFormatInfoSingleton
    {
        [Description("Initializes a new writable instance of the DateTimeFormatInfo class that is culture-independent (invariant).")]
        DateTimeFormatInfo Create();

        // Properties
        DateTimeFormatInfo CurrentInfo
        {
            [Description("Gets a read-only DateTimeFormatInfo object that formats values based on the current culture.")]
            get;
        }

        DateTimeFormatInfo InvariantInfo
        {
            [Description("Gets the default read-only DateTimeFormatInfo object that is culture-independent (invariant).")]
            get;
        }

        // Methods

        [Description("Returns the DateTimeFormatInfo object associated with the specified IFormatProvider.")]
        DateTimeFormatInfo GetInstance(IFormatProvider formatProvider);

        [Description("Returns a read-only DateTimeFormatInfo wrapper.")]
        DateTimeFormatInfo ReadOnly(DateTimeFormatInfo dtfi);
    }
}
