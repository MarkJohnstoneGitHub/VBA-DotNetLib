using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("46470707-6C06-445E-8FAA-91B0BB4927EE")]
    [Description("NumberFormatInfo factory methods and static members that provide culture-specific information for formatting and parsing numeric values.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface INumberFormatInfoSingleton
    {
        [Description("Initializes a new writable instance of the NumberFormatInfo class that is culture-independent (invariant).")]
        NumberFormatInfo Create();

        // Properties
        NumberFormatInfo CurrentInfo
        {
            [Description("Gets a read-only NumberFormatInfo that formats values based on the current culture.")]
            get;
        }

        NumberFormatInfo InvariantInfo 
        {
            [Description("Gets a read-only NumberFormatInfo object that is culture-independent (invariant).")]
            get;
        }

        // Methods
        [Description("Gets the NumberFormatInfo associated with the specified IFormatProvider.")]
        NumberFormatInfo GetInstance(IFormatProvider formatProvider);

        [Description("Returns a read-only NumberFormatInfo wrapper.")]
        NumberFormatInfo ReadOnly(NumberFormatInfo nfi);
    }
}
