// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.numberformatinfo?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using GGlobalization = global::System.Globalization;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("475108FC-A0C6-4E9B-8B95-D03BF61A2C98")]
    [Description("Provides culture-specific information for formatting and parsing numeric values.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface INumberFormatInfo
    {

        // Properties
        int CurrencyDecimalDigits 
        {
            [Description("Gets or sets the number of decimal places to use in currency values.")]
            get;
            [Description("Gets or sets the number of decimal places to use in currency values.")]
            set; 
        }

        string CurrencyDecimalSeparator
        {
            [Description("Gets or sets the string to use as the decimal separator in currency values.")]
            get;
            [Description("Gets or sets the string to use as the decimal separator in currency values.")]
            set;
        }

        string CurrencyGroupSeparator 
        {
            [Description("Gets or sets the string that separates groups of digits to the left of the decimal in currency values.")]
            get;
            [Description("Gets or sets the string that separates groups of digits to the left of the decimal in currency values.")]
            set;
        }

        int[] CurrencyGroupSizes 
        {
            [Description("Gets or sets the number of digits in each group to the left of the decimal in currency values.")]
            get;
            [Description("Gets or sets the number of digits in each group to the left of the decimal in currency values.")]
            set; 
        }

        int CurrencyNegativePattern 
        {
            [Description("Gets or sets the format pattern for negative currency values.")]
            get;
            [Description("Gets or sets the format pattern for negative currency values.")]
            set;
        }

        int CurrencyPositivePattern 
        {
            [Description("Gets or sets the format pattern for positive currency values.")]
            get;
            [Description("Gets or sets the format pattern for positive currency values.")]
            set;
        }

        string CurrencySymbol 
        {
            [Description("Gets or sets the string to use as the currency symbol.")]
            get;
            [Description("Gets or sets the string to use as the currency symbol.")]
            set;
        }

        GGlobalization.DigitShapes DigitSubstitution 
        {
            [Description("Gets or sets a value that specifies how the graphical user interface displays the shape of a digit")]
            get;
            [Description("Gets or sets a value that specifies how the graphical user interface displays the shape of a digit")]
            set; 
        }

        bool IsReadOnly 
        {
            [Description("Gets a value that indicates whether this NumberFormatInfo object is read-only.")]
            get;
        }

        string NaNSymbol 
        {
            [Description("Gets or sets the string that represents the IEEE NaN (not a number) value.")]
            get;
            [Description("Gets or sets the string that represents the IEEE NaN (not a number) value.")]
            set;
        }

        string[] NativeDigits 
        {
            [Description("Gets or sets a string array of native digits equivalent to the Western digits 0 through 9.")]
            get;
            [Description("Gets or sets a string array of native digits equivalent to the Western digits 0 through 9.")]
            set;
        }

        string NegativeInfinitySymbol 
        {
            [Description("Gets or sets the string that represents negative infinity.")]
            get;
            [Description("Gets or sets the string that represents negative infinity.")]
            set;
        }

        string NegativeSign 
        {
            [Description("Gets or sets the string that denotes that the associated number is negative.")]
            get;
            [Description("Gets or sets the string that denotes that the associated number is negative.")]
            set;
        }

        int NumberDecimalDigits 
        {
            [Description("Gets or sets the number of decimal places to use in numeric values.")]
            get;
            [Description("Gets or sets the number of decimal places to use in numeric values.")]
            set;
        }

        string NumberDecimalSeparator
        {
            [Description("Gets or sets the string to use as the decimal separator in numeric values.")]
            get;
            [Description("Gets or sets the string to use as the decimal separator in numeric values.")]
            set;
        }

        string NumberGroupSeparator 
        {
            [Description("Gets or sets the string that separates groups of digits to the left of the decimal in numeric values.")]
            get;
            [Description("Gets or sets the string that separates groups of digits to the left of the decimal in numeric values.")]
            set;
        }

        int[] NumberGroupSizes 
        {
            [Description("Gets or sets the number of digits in each group to the left of the decimal in numeric values.")]
            get;
            [Description("Gets or sets the number of digits in each group to the left of the decimal in numeric values.")]
            set;
        }

        int NumberNegativePattern 
        {
            [Description("Gets or sets the format pattern for negative numeric values.")]
            get;
            [Description("Gets or sets the format pattern for negative numeric values.")]
            set;
        }

        int PercentDecimalDigits 
        {
            [Description("Gets or sets the number of decimal places to use in percent values.")]
            get;
            [Description("Gets or sets the number of decimal places to use in percent values.")]
            set;
        }

        string PercentDecimalSeparator 
        {
            [Description("Gets or sets the string to use as the decimal separator in percent values.")]
            get;
            [Description("Gets or sets the string to use as the decimal separator in percent values.")]
            set;
        }

        string PercentGroupSeparator 
        {
            [Description("Gets or sets the string that separates groups of digits to the left of the decimal in percent values.")]
            get;
            [Description("Gets or sets the string that separates groups of digits to the left of the decimal in percent values.")]
            set;
        }

        int[] PercentGroupSizes 
        {
            [Description("Gets or sets the number of digits in each group to the left of the decimal in percent values.")]
            get;
            [Description("Gets or sets the number of digits in each group to the left of the decimal in percent values.")]
            set;
        }

        int PercentNegativePattern 
        {
            [Description("Gets or sets the format pattern for negative percent values.")]
            get;
            [Description("Gets or sets the format pattern for negative percent values.")]
            set;
        }

        int PercentPositivePattern 
        {
            [Description("Gets or sets the format pattern for positive percent values.")]
            get;
            [Description("Gets or sets the format pattern for positive percent values.")]
            set;
        }

        string PercentSymbol 
        {
            [Description("Gets or sets the string to use as the percent symbol.")]
            get;
            [Description("Gets or sets the string to use as the percent symbol.")]
            set;
        }

        string PerMilleSymbol 
        {
            [Description("Gets or sets the string to use as the per mille symbol.")]
            get;
            [Description("Gets or sets the string to use as the per mille symbol.")]
            set;
        }

        string PositiveInfinitySymbol 
        {
            [Description("Gets or sets the string that represents positive infinity.")]
            get;
            [Description("Gets or sets the string that represents positive infinity.")]
            set;
        }

        string PositiveSign 
        {
            [Description("Gets or sets the string that denotes that the associated number is positive.")]
            get;
            [Description("Gets or sets the string that denotes that the associated number is positive.")]
            set;
        }

        // Methods

        [Description("Creates a shallow copy of the NumberFormatInfo object.")]
        object Clone();

        [Description("Gets an object of the specified type that provides a number formatting service.")]
        object GetFormat(Type formatType);

    }
}
