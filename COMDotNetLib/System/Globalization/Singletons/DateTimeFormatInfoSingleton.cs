using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("A645016B-03D7-4F0E-9CB3-2F63D2D40A8C")]
    [ProgId("DotNetLib.System.Globalization.DateTimeFormatInfoSingleton")]
    [Description("DateTimeFormatInfo factory methods and static members that provides culture-specific information about the format of date and time values.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDateTimeFormatInfoSingleton))]
    public class DateTimeFormatInfoSingleton : IDateTimeFormatInfoSingleton
    {
        public DateTimeFormatInfoSingleton() { }

        // Factory methods
        public DateTimeFormatInfo Create()
        {
            return new DateTimeFormatInfo();
        }

        // Properties
        public DateTimeFormatInfo CurrentInfo => DateTimeFormatInfo.CurrentInfo;

        public DateTimeFormatInfo InvariantInfo => DateTimeFormatInfo.InvariantInfo;

        // Methods
        public DateTimeFormatInfo GetInstance(IFormatProvider formatProvider)
        {
            return DateTimeFormatInfo.GetInstance(formatProvider);
        }

        public DateTimeFormatInfo ReadOnly(DateTimeFormatInfo dtfi)
        {
            return DateTimeFormatInfo.ReadOnly(dtfi);
        }

    }
}
