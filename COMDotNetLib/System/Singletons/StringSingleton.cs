// https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

using GGlobalization = System.Globalization;
using DotNetLib.Extensions;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using CultureInfo = DotNetLib.System.Globalization.CultureInfo;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Description("Represents text as a sequence of UTF-16 code units.")]
    [Guid("2BB0ED15-8B6E-4D70-9DA2-A1C1BA9F8EC3")]
    [ProgId("DotNetLib.System.StringSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IStringSingleton))]

    public class StringSingleton : IStringSingleton
    {
        public StringSingleton() { }


        // Fields
        public string EmptyString => string.Empty;


        //  Compare Overloads

        //public int Compare(string strA, string strB)
        //{
        //    return string.Compare(strA, strB);
        //}

        public int Compare(string strA, string strB, bool ignoreCase = false)
        {
            return string.Compare(strA, strB, ignoreCase);
        }

        public int Compare2(string strA, string strB, StringComparison comparisonType)
        {
            return string.Compare(strA, strB, comparisonType);
        }

        public int Compare3(string strA, string strB, bool ignoreCase, CultureInfo culture)
        {
            return string.Compare(strA, strB, ignoreCase, (GGlobalization.CultureInfo)culture.Unwrap());
        }

        public int Compare4(string strA, string strB, CultureInfo culture, CompareOptions options)
        {
            return string.Compare(strA, strB, (GGlobalization.CultureInfo)culture.Unwrap(), options);
        }

        //public int Compare(string strA, int indexA, string strB, int indexB, int length)
        //{
        //    return string.Compare(strA, indexA, strB, indexB, length);
        //}

        public int Compare5(string strA, int indexA, string strB, int indexB, int length, bool ignoreCase = false)
        {
            return string.Compare(strA, indexA, strB, indexB, length, ignoreCase);
        }

        public int Compare6(string strA, int indexA, string strB, int indexB, int length, StringComparison comparisonType)
        {
            return string.Compare(strA, indexA, strB, indexB, length, comparisonType);
        }

        public int Compare7(string strA, int indexA, string strB, int indexB, int length, bool ignoreCase, CultureInfo culture)
        {
            return string.Compare(strA, indexA, strB, indexB, length, ignoreCase, (GGlobalization.CultureInfo)culture.Unwrap());
        }

        public int Compare8(string strA, int indexA, string strB, int indexB, int length, CultureInfo culture, CompareOptions options)
        {
            return string.Compare(strA, indexA, strB, indexB, length, (GGlobalization.CultureInfo)culture.Unwrap(), options);
        }

        // CompareOrdinal Overloads

        public int CompareOrdinal(string strA, string strB)
        {
            return string.CompareOrdinal(strA, strB);
        }

        public int CompareOrdinal2(string strA, int indexA, string strB, int indexB, int length)
        {
            return string.CompareOrdinal(strA, indexA, strB, indexB, length);
        }

        public string Copy(string str)
        {
            return string.Copy(str);
        }

        public bool Equals(string a, string b)
        {
            return a == b;
        }
        public bool Equals2(string a, string b, StringComparison comparisonType)
        {
            return string.Equals(a, b, comparisonType);
        }

        // Format Overloads

        public string Format(string pFormat, [In] ref object[] args)
        {
            return string.Format(pFormat, args.Unwrap());
        }
        public string Format2(IFormatProvider provider, string pFormat, [In] ref object[] args)
        {
            return string.Format(provider.Unwrap(), pFormat, args.Unwrap());
        }

        public bool IsNullOrEmpty(string value)
        {
            return string.IsNullOrEmpty(value);
        }

        public bool IsNullOrWhiteSpace(string value)
        {
            return string.IsNullOrWhiteSpace(value);
        }
    }
}


//public string Format(string format, object arg0)
//{  
//    return string.Format(format, arg0.Unwrap()); 
//}

//public string Format(string format, object arg0, object arg1)
//{
//    return string.Format(format, arg0.Unwrap(), arg1.Unwrap());
//}

//public string Format(string format, object arg0, object arg1, object arg2)
//{
//    return string.Format(format, arg0.Unwrap(), arg1.Unwrap(), arg2.Unwrap());
//}

//public string Format(string format, object arg0, object arg1, object arg2, object arg3)
//{
//    return string.Format(format, arg0.Unwrap(), arg1.Unwrap(), arg2.Unwrap(), arg3.Unwrap());
//}

//public string Format(IFormatProvider provider, string format, object arg0)
//{
//    return string.Format(provider, format, arg0.Unwrap());
//}

//public static string Format(IFormatProvider provider, string format, object arg0, object arg1)
//{
//    return string.Format(provider, format, arg0.Unwrap(), arg1.Unwrap());
//}

//public string Format(IFormatProvider provider, string format, object arg0, object arg1, object arg2)
//{
//    return string.Format(provider, format, arg0.Unwrap(), arg1.Unwrap(), arg2.Unwrap());
//}