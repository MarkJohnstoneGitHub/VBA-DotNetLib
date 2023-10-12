// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.numberformatinfo?view=netframework-4.8.1
using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using GGlobalization = global::System.Globalization;

namespace DotNetLib.System.Globalization
{
    [Serializable]
    [ComVisible(true)]
    [Guid("18B2AE13-8066-4A94-A770-A90EBA1C6487")]
    [ProgId("DotNetLib.System.Globalization.NumberFormatInfo")]
    [Description("Provides culture-specific information for formatting and parsing numeric values.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(INumberFormatInfo))]
    public class NumberFormatInfo : ICloneable, IFormatProvider, INumberFormatInfo
    {
        private GGlobalization.NumberFormatInfo _numberFormatInfo;
        private Func<object> clone;

        public NumberFormatInfo() 
        {
            _numberFormatInfo = new GGlobalization.NumberFormatInfo();
        }

        public NumberFormatInfo(GGlobalization.NumberFormatInfo numberFormatInfo)
        {
            _numberFormatInfo = numberFormatInfo;
        }

        public NumberFormatInfo(Func<object> clone)
        {
            this.clone = clone;
        }

        //public NumberFormatInfo(object v)
        //{
        //}

        // Properties
        public GGlobalization.NumberFormatInfo WrappedNumberFormatInfo
        {
            get { return this._numberFormatInfo; }
            set { this._numberFormatInfo = value; }
        }

        public int CurrencyDecimalDigits 
        { 
            get => _numberFormatInfo.CurrencyDecimalDigits;
            set => _numberFormatInfo.CurrencyDecimalDigits = value;
        }
        
        public string CurrencyDecimalSeparator 
        { 
            get => _numberFormatInfo.CurrencyDecimalSeparator;
            set => _numberFormatInfo.CurrencyDecimalSeparator = value; 
        }

        public string CurrencyGroupSeparator 
        {
            get => _numberFormatInfo.CurrencyGroupSeparator;
            set => _numberFormatInfo.CurrencyGroupSeparator = value;
        }
        public int[] CurrencyGroupSizes 
        {
            get => _numberFormatInfo.CurrencyGroupSizes;
            set => _numberFormatInfo.CurrencyGroupSizes = value;
        }

        public int CurrencyNegativePattern 
        {
            get => _numberFormatInfo.CurrencyNegativePattern;
            set => _numberFormatInfo.CurrencyNegativePattern = value;
        }
        public int CurrencyPositivePattern 
        { 
            get => _numberFormatInfo.CurrencyPositivePattern;
            set => _numberFormatInfo.CurrencyPositivePattern = value;
        }
        public string CurrencySymbol 
        { 
            get => _numberFormatInfo.CurrencySymbol; 
            set => _numberFormatInfo.CurrencySymbol = value;
        }

        public static NumberFormatInfo CurrentInfo 
        { 
            get => new NumberFormatInfo(GGlobalization.NumberFormatInfo.CurrentInfo);
        }
        public DigitShapes DigitSubstitution 
        { 
            get => _numberFormatInfo.DigitSubstitution;
            set => _numberFormatInfo.DigitSubstitution = value;
        }

        public static NumberFormatInfo InvariantInfo 
        { 
            get => new NumberFormatInfo(GGlobalization.NumberFormatInfo.InvariantInfo);
        }

        public bool IsReadOnly => _numberFormatInfo.IsReadOnly;

        public string NaNSymbol 
        { 
            get => _numberFormatInfo.NaNSymbol;
            set => _numberFormatInfo.NaNSymbol = value;
        }
        public string[] NativeDigits 
        { 
            get => _numberFormatInfo.NativeDigits;
            set => _numberFormatInfo.NativeDigits = value;
        }

        public string NegativeInfinitySymbol 
        { 
            get => _numberFormatInfo.NegativeInfinitySymbol;
            set => _numberFormatInfo.NegativeInfinitySymbol = value;
        }

        public string NegativeSign 
        { 
            get => _numberFormatInfo.NegativeSign;
            set => _numberFormatInfo.NegativeSign = value;
        }

        public int NumberDecimalDigits 
        { 
            get => _numberFormatInfo.NumberDecimalDigits;
            set => _numberFormatInfo.NumberDecimalDigits = value;
        }

        public string NumberDecimalSeparator 
        { 
            get => _numberFormatInfo.NumberDecimalSeparator;
            set => _numberFormatInfo.NumberDecimalSeparator = value;
        }

        public string NumberGroupSeparator 
        {
            get => _numberFormatInfo.NumberGroupSeparator;
            set => _numberFormatInfo.NumberGroupSeparator= value;
        }

        public int[] NumberGroupSizes 
        { 
            get => _numberFormatInfo.NumberGroupSizes;
            set => _numberFormatInfo.NumberGroupSizes = value;
        }

        public int NumberNegativePattern 
        { 
            get => _numberFormatInfo.NumberNegativePattern; 
            set => _numberFormatInfo.NumberNegativePattern = value; 
        }

        public int PercentDecimalDigits 
        { 
            get => _numberFormatInfo.PercentDecimalDigits;
            set => _numberFormatInfo.PercentDecimalDigits = value;
        }
        public string PercentDecimalSeparator 
        { 
            get => _numberFormatInfo.PercentDecimalSeparator; 
            set => _numberFormatInfo.PercentDecimalSeparator = value;
        }
        public string PercentGroupSeparator 
        { 
            get => _numberFormatInfo.PercentGroupSeparator;
            set => _numberFormatInfo.PercentGroupSeparator = value;
        }

        public int[] PercentGroupSizes 
        { 
            get => _numberFormatInfo.PercentGroupSizes;
            set => _numberFormatInfo.PercentGroupSizes = value;
        }

        public int PercentNegativePattern 
        { 
            get => _numberFormatInfo.PercentNegativePattern;
            set => _numberFormatInfo.PercentNegativePattern = value;
        }
        public int PercentPositivePattern 
        { 
            get => _numberFormatInfo.PercentPositivePattern; 
            set => _numberFormatInfo.PercentPositivePattern = value;
        }
        public string PercentSymbol 
        { 
            get => _numberFormatInfo.PercentSymbol;
            set => _numberFormatInfo.PercentSymbol = value;
        }

        public string PerMilleSymbol 
        { 
            get => _numberFormatInfo.PerMilleSymbol; 
            set => _numberFormatInfo.PerMilleSymbol = value;
        }
        public string PositiveInfinitySymbol 
        { 
            get => _numberFormatInfo.PositiveInfinitySymbol; 
            set => _numberFormatInfo.PositiveInfinitySymbol = value;
        }

        public string PositiveSign 
        { 
            get => _numberFormatInfo.PositiveSign;
            set => _numberFormatInfo.PositiveSign = value;
        }

        // Methods
        public object Clone()
        {
            return new NumberFormatInfo(_numberFormatInfo.Clone);
        }

        // TODO: Check implementation
        public object GetFormat(GSystem.Type formatType)
        {
            if (!(formatType == typeof(NumberFormatInfo)))
            {
                return null;
            }
            return this;
        }

        public object GetFormat(Type formatType)
        {
            if (!(formatType.WrappedType == typeof(NumberFormatInfo)))
            {
                return null;
            }
            return this;
        }

        public static NumberFormatInfo GetInstance(IFormatProvider formatProvider)
        {
            return new NumberFormatInfo(GGlobalization.NumberFormatInfo.GetInstance(formatProvider));
        }

        public static NumberFormatInfo ReadOnly(NumberFormatInfo nfi)
        { 
            return new NumberFormatInfo(GGlobalization.NumberFormatInfo.ReadOnly(nfi.WrappedNumberFormatInfo)); 
        }
    }
}
