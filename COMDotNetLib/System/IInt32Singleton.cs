
using GSystem = global::System;
using GGlobalization = global::System.Globalization;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("09AB7234-C897-4E09-A3C8-85BE61F9FDC0")]
    [Description("Represents a 32-bit signed integer.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IInt32Singleton
    {
        //Fields
        int MaxValue
        {
            [Description("Represents the largest possible value of an Int32. This field is constant.")]
            get;
        }

        int MinValue
        {
            [Description("Represents the smallest possible value of Int32. This field is constant.")]
            get;
        }

        // Methods

        [Description("Converts the string representation of a number to its 32-bit signed integer equivalent. A return value indicates whether the conversion succeeded.")]
        bool TryParse(string s, out int result);

        [Description("Converts the string representation of a number in a specified style and culture-specific format to its 32-bit signed integer equivalent. A return value indicates whether the conversion succeeded.")]
        bool TryParse(string s, GGlobalization.NumberStyles style, GSystem.IFormatProvider provider, out int result);

    }
}
