// https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

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

        public bool Contains(string str, string substring, GSystem.StringComparison comparisonType = GSystem.StringComparison.Ordinal)
        {
            return str.Contains(substring, comparisonType);
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
