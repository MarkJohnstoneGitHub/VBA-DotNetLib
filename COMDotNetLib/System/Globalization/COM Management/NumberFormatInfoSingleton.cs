// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.numberformatinfo?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("BA64C728-9E84-4F04-B18C-7773DF6D0E11")]
    [ProgId("DotNetLib.System.Globalization.NumberFormatInfoSingleton")]
    [Description("NumberFormatInfo factory methods and static members that provide culture-specific information for formatting and parsing numeric values.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(INumberFormatInfoSingleton))]
    public class NumberFormatInfoSingleton : INumberFormatInfoSingleton
     {
        public NumberFormatInfoSingleton() { }

        // Factory methods
        public NumberFormatInfo Create()
        { 
            return new NumberFormatInfo();
        }

        // Properties
        public NumberFormatInfo CurrentInfo
        {
            get => NumberFormatInfo.CurrentInfo;
        }

        public NumberFormatInfo InvariantInfo
        {
            get => NumberFormatInfo.InvariantInfo;
        }

        // Methods
        public NumberFormatInfo GetInstance(IFormatProvider formatProvider)
        {
            return NumberFormatInfo.GetInstance(formatProvider);
        }

        public NumberFormatInfo ReadOnly(NumberFormatInfo nfi)
        {
            return NumberFormatInfo.ReadOnly(nfi);
        }

    }


}
