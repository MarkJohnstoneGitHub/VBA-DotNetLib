// https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

using GGeneric = global::System.Collections.Generic;
using GRegularExpressions = global::System.Text.RegularExpressions;
using GSystem = global::System;
using GCollections = global::System.Collections;
using GGlobalization = System.Globalization;
using CultureInfo = DotNetLib.System.Globalization.CultureInfo;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using DotNetLib.Extensions;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Description("Represents text as a sequence of UTF-16 code units.")]
    [Guid("3EE0B1C3-0726-49F4-8E81-56ECA433EC9A")]
    [ProgId("DotNetLib.System.BStringSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IBStringSingleton))]

    public class BStringSingleton : IBStringSingleton
    {
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

        public int CompareOrdinal(string strA, string strB)
        {
            return string.CompareOrdinal(strA, strB);
        }

        public int CompareOrdinal(string strA, int indexA, string strB, int indexB, int length)
        {
            return string.CompareOrdinal(strA, indexA, strB, indexB, length);
        }

        public string Concat([In] ref string[] values)
        {
            return string.Concat(values);
        }

        public string Concat2(string str0, string str1)
        {
            return string.Concat(str0, str1);
        }

        public string Concat3(string str0, string str1, string str2)
        {
            return string.Concat(str0, str1, str2);
        }

        public string Concat4(string str0, string str1, string str2, string str3)
        {
            return string.Concat(str0, str1, str2, str3);
        }

        public string Concat5(GCollections.IEnumerable stringValues)
        {
            return string.Concat((GGeneric.IEnumerable<string>)stringValues);
        }

        public string Concat6([In] ref object[] values)
        {
            return string.Concat(values);
        }

        public string Concat7(object arg0)
        {
            return string.Concat(arg0);
        }

        public string Concat8(object arg0, object arg1)
        {
            return string.Concat(arg0, arg1);
        }

        public string Concat9(object arg0, object arg1, object arg2)
        {
            return string.Concat(arg0, arg1, arg2);
        }

        //public static string Concat<T>(System.Collections.Generic.IEnumerable<T> values);
        public string Concat10(GCollections.IEnumerable values)
        {
            return string.Concat((GGeneric.IEnumerable<object>)values);
        }

        public bool Contains(string str, string substring, GSystem.StringComparison comparisonType = GSystem.StringComparison.Ordinal)
        {
            return str.IndexOf(substring, comparisonType) >= 0;

            //return str.Contains(substring, comparisonType);
        }

        public string Copy(string str)
        {
            return string.Copy(str);
        }

        public bool EndsWith(string str, string substring, GSystem.StringComparison comparisonType = StringComparison.CurrentCulture)
        {
            return str.EndsWith(substring, comparisonType);
        }

        public bool EndsWith2(string str, string substring, bool ignoreCase, CultureInfo culture)
        {
            return str.EndsWith(substring, ignoreCase, culture.WrappedCultureInfo);
        }

        public bool Equals(string a, string b)
        {
            return a == b;
        }

        public bool Equals2(string a, string b, StringComparison comparisonType)
        {
            return string.Equals(a, b, comparisonType);
        }

        public string Format(string pFormat, [In] ref object[] args)
        {
            return string.Format(pFormat, args.Unwrap());
        }

        public string Format2(string pFormat, object arg0)
        {
            return string.Format(pFormat, arg0.Unwrap());
        }

        public string Format3(string pFormat, object arg0, object arg1)
        {
            return string.Format(pFormat, arg0.Unwrap(), arg1.Unwrap());
        }

        public string Format4(string pFormat, object arg0, object arg1, object arg2)
        {
            return string.Format(pFormat, arg0.Unwrap(), arg1.Unwrap(), arg2.Unwrap());
        }

        public string Format5(IFormatProvider provider, string pFormat, [In] ref object[] args)
        {
            return string.Format(provider.Unwrap(), pFormat, args.Unwrap());
        }

        public string Format6(IFormatProvider provider, string pFormat, object arg0)
        {
            return string.Format(provider.Unwrap(), pFormat, arg0.Unwrap());
        }

        public string Format7(IFormatProvider provider, string pFormat, object arg0, object arg1)
        {
            return string.Format(provider.Unwrap(), pFormat, arg0.Unwrap(), arg1.Unwrap());
        }

        public string Format8(IFormatProvider provider, string pFormat, object arg0, object arg1, object arg2)
        {
            return string.Format(provider.Unwrap(), pFormat, arg0.Unwrap(), arg1.Unwrap(), arg2.Unwrap());
        }

        public bool IsNullOrEmpty(string value)
        {
            return string.IsNullOrEmpty(value);
        }
     
        public bool IsNullOrWhiteSpace(string value)
        {
            return string.IsNullOrWhiteSpace(value);
        }

        public string Join(string separator, [In] ref string[] value)
        {
            return string.Join(separator, value);
        }

        public string Join2(string separator, [In] ref object[] values)
        {
            return string.Join(separator, values.Unwrap());
        }

        public string Join3(string separator, GCollections.IEnumerable stringValues)
        {
            return string.Join(separator, (IEnumerable<string>)stringValues);
        }

        public string Join4(string separator, [In] ref string[] value, int startIndex, int count)
        {
            return string.Join(separator, value, startIndex, count);
        }

        public string[] Split(string str, string separators, StringSplitOptions options = StringSplitOptions.None)
        {
            if (separators == null)
            {
                return str.Split(null as char[], (GSystem.StringSplitOptions)options);
            }
            return str.Split(separators.ToCharArray(), (GSystem.StringSplitOptions)options);
        }

        public string[] Split2(string str, string separator, int count, StringSplitOptions options = StringSplitOptions.None)
        {
            if (separator == null)
            {
                return str.Split(null as char[], count, (GSystem.StringSplitOptions)options);
            }

            return str.Split(separator.ToCharArray(), count, (GSystem.StringSplitOptions)options);
        }

        public string[] Split3(string str, [In] ref string[] separator, StringSplitOptions options)
        {
            return str.Split(separator, (GSystem.StringSplitOptions)options);
        }

        public string[] Split4(string str, [In] ref string[] separator, int count, StringSplitOptions options)
        {
            return str.Split(separator, count, (GSystem.StringSplitOptions)options);
        }

        public bool StartsWith(string str, string substring, GSystem.StringComparison comparisonType = StringComparison.CurrentCulture)
        {
            return str.StartsWith(substring, comparisonType);
        }

        public bool StartsWith2(string str, string substring, bool ignoreCase, CultureInfo culture)
        {
            return str.StartsWith(substring, ignoreCase, culture.WrappedCultureInfo);
        }

        public string Unescape(string value)
        {
            if (value == null)
                return string.Empty;
            return GRegularExpressions.Regex.Unescape(value);
        }

    }
}
