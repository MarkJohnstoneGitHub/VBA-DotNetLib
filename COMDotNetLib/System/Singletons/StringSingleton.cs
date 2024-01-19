// https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

//Todo for constructors if value parameter is null return null or empty string?

using DotNetLib.Extensions;
using GGeneric = global::System.Collections.Generic;
using GSystem = global::System;
using GRegularExpressions = global::System.Text.RegularExpressions;
using GCollections = global::System.Collections;
using GGlobalization = System.Globalization;
using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using CultureInfo = DotNetLib.System.Globalization.CultureInfo;
using System.Collections.Generic;

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

        // Factory Methods

        // Notes https://stackoverflow.com/questions/25759878/convert-byte-to-sbyte 
        // Possible add String(SByte*, Int32, Int32) by passing an array of bytes? 
        public String Create(string value)
        {
            if (value == null)
                return null;
            return new String(value); 
        }

        public String Create(string character, int count)
        {
            if (character == null)
                return null;
            return new String(character, count);
        }

        public String Create(string value, int startIndex, int length)
        {
            if (value == null)
                return null;
            return new String(value, startIndex, length);
        }

        public String Create(String value)
        {
            if (value == null)
                return null;
            return new String(value);
        }

        public String Create(String character, int count)
        {
            if (character == null || character.WrappedString == null)
                return null;
            return new String(character, count);
        }

        public String Create(String value, int startIndex, int length)
        {
            if (value == null || value.WrappedString == null)
                return null;
            return new String(value, startIndex, length);
        }

        public String CreateUnescape(string value)
        {
            if (value == null)
                return null;
            return new String(GRegularExpressions.Regex.Unescape(value)); 
        }

        // Fields

        public String EmptyString => String.Empty;

        //  Methods

        public int Compare(String strA, String strB, bool ignoreCase = false)
        {
            return string.Compare(strA.WrappedString, strB.WrappedString, ignoreCase);
        }

        public int Compare2(String strA, String strB, StringComparison comparisonType)
        {
            return string.Compare(strA.WrappedString, strB.WrappedString, comparisonType);
        }

        public int Compare3(String strA, String strB, bool ignoreCase, CultureInfo culture)
        {
            return string.Compare(strA.WrappedString, strB.WrappedString, ignoreCase, (GGlobalization.CultureInfo)culture.Unwrap());
        }

        public int Compare4(String strA, String strB, CultureInfo culture, CompareOptions options)
        {
            return string.Compare(strA.WrappedString, strB.WrappedString, (GGlobalization.CultureInfo)culture.Unwrap(), options);
        }

        public int Compare5(String strA, int indexA, String strB, int indexB, int length, bool ignoreCase = false)
        {
            return string.Compare(strA.WrappedString, indexA, strB.WrappedString, indexB, length, ignoreCase);
        }

        public int Compare6(String strA, int indexA, String strB, int indexB, int length, StringComparison comparisonType)
        {
            return string.Compare(strA.WrappedString, indexA, strB.WrappedString, indexB, length, comparisonType);
        }

        public int Compare7(String strA, int indexA, String strB, int indexB, int length, bool ignoreCase, CultureInfo culture)
        {
            return string.Compare(strA.WrappedString, indexA, strB.WrappedString, indexB, length, ignoreCase, (GGlobalization.CultureInfo)culture.Unwrap());
        }

        public int Compare8(String strA, int indexA, String strB, int indexB, int length, CultureInfo culture, CompareOptions options)
        {
            return string.Compare(strA.WrappedString, indexA, strB.WrappedString, indexB, length, (GGlobalization.CultureInfo)culture.Unwrap(), options);
        }

        public int CompareOrdinal(String strA, String strB)
        {
            return string.CompareOrdinal(strA.WrappedString, strB.WrappedString);
        }

        public int CompareOrdinal2(String strA, int indexA, String strB, int indexB, int length)
        {
            return string.CompareOrdinal(strA.WrappedString, indexA, strB.WrappedString, indexB, length);
        }

        public String Concat(object arg0)
        {  
            return new String(string.Concat(arg0));
        }

        public String Concat2(object arg0, object arg1)
        {
            return new String(string.Concat(arg0, arg1));
        }

        public String Concat3(object arg0, object arg1, object arg2)
        {
            return new String(string.Concat(arg0, arg1, arg2));
        }

        public  String Concat4([In] ref object[] values)
        {
            return new String(string.Concat(values));
        }

        public String Concat5(GCollections.IEnumerable stringValues)
        {
            return new String(string.Concat((GGeneric.IEnumerable<string>)stringValues));
        }

        //public static string Concat8<T>(System.Collections.Generic.IEnumerable<T> values);
        public String Concat6(GCollections.IEnumerable values)
        {
            return new String(string.Concat((GGeneric.IEnumerable<object>)values));
        }

        public String Concat7(string str0, string str1)
        {
            return new String(string.Concat(str0, str1));
        }

        public String Concat8(string str0, string str1, string str2)
        {
            return new String(string.Concat(str0, str1, str2));
        }

        public String Concat9(string str0, string str1, string str2, string str3)
        {
            return new String(string.Concat(str0, str1, str2, str3));
        }

        public String Concat10([In] ref string[] values)
        {
            return new String(string.Concat(values));
        }

        public String Concat12(String str0, String str1)
        {
            return new String(string.Concat(str0.WrappedString, str1.WrappedString));
        }

        public String Concat13(String str0, String str1, String str2)
        {
            return new String(string.Concat(str0.WrappedString, str1.WrappedString, str2.WrappedString));
        }

        public String Concat14(String str0, String str1, String str2, String str3)
        {
            return new String(string.Concat(str0.WrappedString, str1.WrappedString, str2.WrappedString, str3.WrappedString));
        }


        public String Copy(string str)
        {
            return new String(string.Copy(str));
        }

        public String Copy(String str)
        {
            return new String(string.Copy(str.WrappedString));
        }

        public bool Equals(String a, String b)
        {
            if ((object)a == b)
            {
                return true;
            }

            if ((object)a == null || (object)b == null)
            {
                return false;
            }
            return string.Equals(a.WrappedString, b.WrappedString);
        }

        public bool Equals2(String a, String b, StringComparison comparisonType)
        {
            if ((object)a == b)
            {
                return true;
            }

            if ((object)a == null || (object)b == null)
            {
                return false;
            }
            return string.Equals(a.WrappedString, b.WrappedString, comparisonType);
        }

        public String Format(string pFormat, [In] ref object[] args)
        {
            return new String(string.Format(pFormat, args.Unwrap()));
        }
        public String Format2(IFormatProvider provider, string pFormat, [In] ref object[] args)
        {
            return new String(string.Format(provider.Unwrap(), pFormat, args.Unwrap()));
        }

        public String Format3(String pFormat, [In] ref object[] args)
        {
            return new String(string.Format(pFormat.WrappedString, args.Unwrap()));
        }
        public String Format4(IFormatProvider provider, String pFormat, [In] ref object[] args)
        {
            return new String(string.Format(provider.Unwrap(), pFormat.WrappedString, args.Unwrap()));
        }

        public bool IsNullOrEmpty(String value)
        {
            if ((object)value != null)
            {
                return value.WrappedString.Length == 0; // string.IsNullOrEmpty(value.WrappedString);
            }
            return true;
        }

        public bool IsNullOrWhiteSpace(String value)
        {
            if ((object)value == null)
            {
                return true;
            }
            return string.IsNullOrWhiteSpace(value.WrappedString);
        }

        public String Join(string separator, [In] ref string[] value)
        { 
            return new String(string.Join(separator, value));
        }

        public String Join2(string separator, [In] ref object[] values)
        {
            return new String(string.Join(separator, values.Unwrap()));
        }

        public String Join3(string separator, GCollections.IEnumerable stringValues)
        {
            return new String(string.Join(separator, (IEnumerable<string>)stringValues));
        }

        public String Join4(string separator, [In] ref string[] value, int startIndex, int count)
        {
            return new String(string.Join(separator, value, startIndex, count));
        }

        // Extensions

        public String Unescape(String value)
        {
            if (value == null || value.WrappedString == null)
                return null;
            return new String(GRegularExpressions.Regex.Unescape(value.WrappedString));
        }

        public String Unescape2(string value)
        {
            if (value == null)
                return null;
            return new String(GRegularExpressions.Regex.Unescape(value));
        }

        // Extensions

        public bool IsSurrogate(string str)
        {
            return str.IsSurrogate();
        }

    }
}