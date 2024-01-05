// https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

using GSystem = global::System;
using GText = global::System.Text;
using System;
using DotNetLib.Extensions;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Description("Represents text as a sequence of UTF-16 code units.")]
    [Guid("A1514BB3-A4C7-4B4C-84C8-C12B309AF00F")]
    [ProgId("DotNetLib.System.String")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IString))]
    public class String : IString, ICloneable, IComparable, IWrappedObject
    {
        private GSystem.String _string;

        // Constructors
        public String(string value)
        {
            _string = value;
        }

        public String(string character, int count)
        {
            _string = new string(character[0], count);
        }

        public String(string value, int startIndex, int length)
        {
            _string = new string(value.ToCharArray(), startIndex, length);
        }

        public String(String value)
        {
            
            _string = value.WrappedString;
        }

        public String(String character, int count)
        {
            _string = new string(character.WrappedString[0], count);
        }

        public String(String value, int startIndex, int length)
        {
            _string = new string(value.WrappedString.ToCharArray(), startIndex, length);
        }

        // Properties

        public static readonly String Empty = new String(GSystem.String.Empty);

        public int Length => _string.Length;

        public object WrappedObject => _string;

        public string WrappedString => _string;

        // https://learn.microsoft.com/en-us/dotnet/csharp/programming-guide/indexers/using-indexers
        [IndexerName("CharCodes")]
        public int this[int index] => _string[index];

        // Methods

        //Todo Check implementation
        public object Clone()
        {
            return this;
            //return new String((string)_string.Clone());
        }

        public int CompareTo(object obj)
        {
            return _string.CompareTo(obj.Unwrap());
        }

        public int CompareTo(String strB)
        {
            return _string.CompareTo(strB.WrappedObject);
        }
        public int CompareTo2(string strB)
        {
            return _string.CompareTo(strB);
        }

        public int CompareTo3(object obj)
        {
            return _string.CompareTo(obj.Unwrap());
        }

        public bool Contains(String value, StringComparison comparisonType = StringComparison.Ordinal)
        {
            return _string.IndexOf(value.WrappedString, comparisonType) >= 0;
        }

        public bool Contains2(string value, StringComparison comparisonType = StringComparison.Ordinal)
        {
            return _string.IndexOf(value, comparisonType) >= 0;
        }

        public bool EndsWith(String value, GSystem.StringComparison comparisonType = StringComparison.CurrentCulture)
        {
            return _string.EndsWith(value.WrappedString, comparisonType);
        }

        public bool EndsWith2(String value, bool ignoreCase, CultureInfo culture)
        {
            return _string.EndsWith(value.WrappedString, ignoreCase, culture.WrappedCultureInfo);
        }

        public bool EndsWith3(string value, GSystem.StringComparison comparisonType = StringComparison.CurrentCulture)
        {
            return _string.EndsWith(value, comparisonType);
        }

        public bool EndsWith4(string value, bool ignoreCase, CultureInfo culture)
        {
            return _string.EndsWith(value, ignoreCase, culture.WrappedCultureInfo);
        }

        public override bool Equals(object obj)
        {
            return _string.Equals(obj.Unwrap());
        }

        public bool Equals2(String value)
        {
            return _string.Equals(value.WrappedString);
        }

        public bool Equals3(String value, GSystem.StringComparison comparisonType)
        {
            return _string.Equals(value.WrappedString, comparisonType);
        }

        public bool Equals4(string value)
        {
            return _string.Equals(value);
        }

        public bool Equals5(string value, GSystem.StringComparison comparisonType)
        {
            return _string.Equals(value, comparisonType);
        }

        public override int GetHashCode()
        {
            return _string.GetHashCode();
        }

        public new Type GetType()
        {
            return new Type(((GSystem.Object)this).GetType());
        }

        public GSystem.TypeCode GetTypeCode()
        {
            return GSystem.TypeCode.String;
        }

        public int IndexOf(String value)
        {
            return _string.IndexOf(value.WrappedString);
        }

        public int IndexOf2(String value, GSystem.StringComparison comparisonType)
        {
            return _string.IndexOf(value.WrappedString, comparisonType);
        }

        public int IndexOf3(String value, int startIndex)
        {
            return _string.IndexOf(value.WrappedString, startIndex);
        }

        public int IndexOf4(String value, int startIndex, StringComparison comparisonType)
        {
            return _string.IndexOf(value.WrappedString, startIndex, comparisonType);
        }

        public int IndexOf5(String value, int startIndex, int count)
        {
            return _string.IndexOf(value.WrappedString, startIndex, count);
        }

        public int IndexOf6(String value, int startIndex, int count, GSystem.StringComparison comparisonType)
        {
            return _string.IndexOf(value.WrappedString, startIndex, count, comparisonType);
        }

        public int IndexOf7(string value)
        {
            return _string.IndexOf(value);
        }

        public int IndexOf8(string value, GSystem.StringComparison comparisonType)
        {
            return _string.IndexOf(value, comparisonType);
        }

        public int IndexOf9(string value, int startIndex)
        {
            return _string.IndexOf(value, startIndex);
        }

        public int IndexOf10(string value, int startIndex, StringComparison comparisonType)
        {
            return _string.IndexOf(value, startIndex, comparisonType);
        }

        public int IndexOf11(string value, int startIndex, int count)
        {
            return _string.IndexOf(value, startIndex, count);
        }

        public int IndexOf12(string value, int startIndex, int count, GSystem.StringComparison comparisonType)
        {
            return _string.IndexOf(value, startIndex, count, comparisonType);
        }

        public int IndexOfAny(String anyOf)
        {
            return _string.IndexOfAny(anyOf.WrappedString.ToCharArray());
        }

        public int IndexOfAny2(String anyOf, int startIndex)
        {
            return _string.IndexOfAny(anyOf.WrappedString.ToCharArray(), startIndex);
        }

        public int IndexOfAny3(String anyOf, int startIndex, int count)
        {
            return _string.IndexOfAny(anyOf.WrappedString.ToCharArray(), startIndex, count);
        }

        public int IndexOfAny4(string anyOf)
        {
            return _string.IndexOfAny(anyOf.ToCharArray());
        }

        public int IndexOfAny5(string anyOf, int startIndex)
        {
            return _string.IndexOfAny(anyOf.ToCharArray(), startIndex);
        }

        public int IndexOfAny6(string anyOf, int startIndex, int count)
        {
            return _string.IndexOfAny(anyOf.ToCharArray(), startIndex, count);
        }

        public String Insert(int startIndex, String value)
        {
            return new String(_string.Insert(startIndex, value.WrappedString));
        }

        public String Insert2(int startIndex, string value)
        {
            return new String(_string.Insert(startIndex, value));
        }

        public string InsertBStr(int startIndex, string value)
        {
            return _string.Insert(startIndex, value);
        }

        public bool IsNormalized(GText.NormalizationForm normalizationForm = GText.NormalizationForm.FormC)
        {
            return _string.IsNormalized(normalizationForm);
        }

        public int LastIndexOf(String value)
        {
            return _string.LastIndexOf(value.WrappedString);
        }

        public int LastIndexOf2(String value, int startIndex)
        {
            return _string.LastIndexOf(value.WrappedString, startIndex);
        }

        public int LastIndexOf3(String value, int startIndex, int count)
        {
            return _string.LastIndexOf(value.WrappedString, startIndex, count);
        }

        public int LastIndexOf4(String value, GSystem.StringComparison comparisonType)
        {
            return _string.LastIndexOf(value.WrappedString, comparisonType);
        }

        public int LastIndexOf5(String value, int startIndex, StringComparison comparisonType)
        {
            return _string.LastIndexOf(value.WrappedString, startIndex, comparisonType);
        }

        public int LastIndexOf6(String value, int startIndex, int count, GSystem.StringComparison comparisonType)
        {
            return _string.LastIndexOf(value.WrappedString, startIndex, count, comparisonType);
        }

        public int LastIndexOf7(string value)
        {
            return _string.LastIndexOf(value);
        }

        public int LastIndexOf8(string value, int startIndex)
        {
            return _string.LastIndexOf(value, startIndex);
        }

        public int LastIndexOf9(string value, int startIndex, int count)
        {
            return _string.LastIndexOf(value, startIndex, count);
        }

        public int LastIndexOf10(string value, GSystem.StringComparison comparisonType)
        {
            return _string.LastIndexOf(value, comparisonType);
        }

        public int LastIndexOf11(string value, int startIndex, StringComparison comparisonType)
        {
            return _string.LastIndexOf(value, startIndex, comparisonType);
        }

        public int LastIndexOf12(string value, int startIndex, int count, GSystem.StringComparison comparisonType)
        {
            return _string.LastIndexOf(value, startIndex, count, comparisonType);
        }

        public int LastIndexOfAny(String anyOf)
        {
            return _string.LastIndexOfAny(anyOf.WrappedString.ToCharArray());
        }

        public int LastIndexOfAny2(String anyOf, int startIndex)
        {
            return _string.LastIndexOfAny(anyOf.WrappedString.ToCharArray(), startIndex);
        }

        public int LastIndexOfAny3(String anyOf, int startIndex, int count)
        {
            return _string.LastIndexOfAny(anyOf.WrappedString.ToCharArray(), startIndex, count);
        }

        public int LastIndexOfAny4(string anyOf)
        {
            return _string.LastIndexOfAny(anyOf.ToCharArray());
        }

        public int LastIndexOfAny5(string anyOf, int startIndex)
        {
            return _string.LastIndexOfAny(anyOf.ToCharArray(), startIndex);
        }

        public int LastIndexOfAny6(string anyOf, int startIndex, int count)
        {
            return _string.LastIndexOfAny(anyOf.ToCharArray(), startIndex, count);
        }

        //public String Normalize()
        //{
        //    return new String(_string.Normalize());
        //}

        public String Normalize(GText.NormalizationForm normalizationForm = GText.NormalizationForm.FormC)
        {
            return new String(_string.Normalize(normalizationForm));
        }

        //public string NormalizeBStr()
        //{
        //    return _string.Normalize();
        //}

        public string NormalizeBStr(GText.NormalizationForm normalizationForm = GText.NormalizationForm.FormC)
        {
            return _string.Normalize(normalizationForm);
        }

        public String PadLeft(int totalWidth)
        {
            return new String(_string.PadLeft(totalWidth));
        }

        public String PadLeft(int totalWidth, string paddingChar = " ")
        {
            return new String(_string.PadLeft(totalWidth, paddingChar[0]));
        }

        public string PadLeftBStr(int totalWidth)
        {
            return _string.PadLeft(totalWidth);
        }

        public string PadLeftBStr(int totalWidth, string paddingChar = " ")
        {
            return _string.PadLeft(totalWidth, paddingChar[0]);
        }

        public String PadRight(int totalWidth)
        {
            return new String(_string.PadRight(totalWidth));
        }

        public String PadRight(int totalWidth, string paddingChar = " ")
        {
            return new String(_string.PadRight(totalWidth, paddingChar[0]));
        }

        public string PadRightBStr(int totalWidth)
        {
            return _string.PadRight(totalWidth);
        }

        public string PadRightBStr(int totalWidth, string paddingChar = " ")
        {
            return _string.PadRight(totalWidth, paddingChar[0]);
        }

        public String Remove(int startIndex)
        {
            return new String(_string.Remove(startIndex));
        }

        public String Remove2(int startIndex, int count)
        {
            return new String(_string.Remove(startIndex, count));
        }

        public string RemoveBStr(int startIndex)
        {
            return _string.Remove(startIndex);
        }

        public string RemoveBStr2(int startIndex, int count)
        {
            return _string.Remove(startIndex, count);
        }

        public String Replace(String oldValue, String newValue)
        {
            return new String(_string.Replace(oldValue.WrappedString, newValue.WrappedString));
        }

        public String Replace2(string oldValue, string newValue)
        {
            return new String(_string.Replace(oldValue, newValue));
        }

        public string ReplaceBStr(string oldValue, string newValue)
        {
            return _string.Replace(oldValue, newValue);
        }

        //public string[] Split(string separator)
        //{
        //    return _string.Split(separator.ToCharArray());
        //}

        public string[] Split(string separator, StringSplitOptions options = StringSplitOptions.None)
        {
            return _string.Split(separator.ToCharArray(), (GSystem.StringSplitOptions)options);
        }

        //public string[] Split2(string separator, int count)
        //{
        //    if (separator == null)
        //    {
        //        return _string.Split(null, count);
        //    }
        //    return _string.Split(separator.ToCharArray(), count);
        //}

        public string[] Split2(string separator, int count, StringSplitOptions options = StringSplitOptions.None)
        {
            if (separator == null && options == StringSplitOptions.None)
            {
                return _string.Split(null, count);
            }

            return _string.Split(separator.ToCharArray(), count, (GSystem.StringSplitOptions)options);
        }

        public string[] Split3(string[] separator, StringSplitOptions options)
        {
            return _string.Split(separator, (GSystem.StringSplitOptions)options);
        }

        public string[] Split4(string[] separator, int count, StringSplitOptions options)
        {
            return _string.Split(separator, count, (GSystem.StringSplitOptions)options);
        }

        public Array SplitStringArray(string separator)
        {
            return new Array(_string.Split(separator.ToCharArray()));
        }

        public Array SplitStringArray2(string separator, StringSplitOptions options)
        {
            return new Array(_string.Split(separator.ToCharArray(), (GSystem.StringSplitOptions)options));
        }

        public Array SplitStringArray3(string separator, int count, StringSplitOptions options = StringSplitOptions.None)
        {
            return new Array(_string.Split(separator.ToCharArray(), count, (GSystem.StringSplitOptions)options));
        }

        public Array SplitStringArray4(string[] separator, StringSplitOptions options)
        {
            return new Array(_string.Split(separator, (GSystem.StringSplitOptions)options));
        }

        public Array SplitStringArray5(string[] separator, int count, StringSplitOptions options)
        {
            return new Array(_string.Split(separator, count, (GSystem.StringSplitOptions)options));
        }

        public bool StartsWith(String value, GSystem.StringComparison comparisonType = StringComparison.CurrentCulture)
        {
            return _string.StartsWith(value.WrappedString, comparisonType);
        }

        public bool StartsWith2(String value, bool ignoreCase, CultureInfo culture)
        {
            return _string.StartsWith(value.WrappedString, ignoreCase, culture.WrappedCultureInfo);
        }

        public bool StartsWith3(string value, GSystem.StringComparison comparisonType = StringComparison.CurrentCulture)
        {
            return _string.StartsWith(value, comparisonType);
        }

        public bool StartsWith4(string value, bool ignoreCase, CultureInfo culture)
        {
            return _string.StartsWith(value, ignoreCase, culture.WrappedCultureInfo);
        }

        public String Substring(int startIndex)
        {
            return new String(_string.Substring(startIndex));
        }

        public String Substring2(int startIndex, int length)
        {
            return new String(_string.Substring(startIndex, length));
        }

        public string SubstringBStr(int startIndex)
        {
            return _string.Substring(startIndex);
        }

        public string SubstringBStr2(int startIndex, int length)
        {
            return _string.Substring(startIndex, length);
        }

        public String ToLower()
        {
            return new String(_string.ToLower());
        }

        public String ToLower2(CultureInfo culture)
        {
            return new String(_string.ToLower(culture.WrappedCultureInfo));
        }

        public string ToLowerBStr()
        {
            return _string.ToLower();
        }

        public string ToLowerBStr2(CultureInfo culture)
        {
            return _string.ToLower(culture.WrappedCultureInfo);
        }

        public String ToLowerInvariant()
        {
            return new String(_string.ToLowerInvariant());
        }

        public string ToLowerInvariantBStr()
        {
            return _string.ToLowerInvariant();
        }

        public override string ToString()
        {
            return _string;
        }

        public String ToUpper()
        {
            return new String(_string.ToUpper());
        }

        public String ToUpper2(CultureInfo culture)
        {
            return new String(_string.ToUpper(culture.WrappedCultureInfo));
        }

        public string ToUpperBStr()
        {
            return _string.ToUpper();
        }

        public string ToUpperBStr2(CultureInfo culture)
        {
            return _string.ToUpper(culture.WrappedCultureInfo);
        }

        public String ToUpperInvariant()
        {
            return new String(_string.ToUpperInvariant());
        }

        public string ToUpperInvariantBStr()
        { 
            return _string.ToUpperInvariant(); 
        }

        public String Trim()
        {
            return new String(_string.Trim());
        }
        public String Trim2(String trimChars)
        {
            return new String(_string.Trim(trimChars.WrappedString.ToCharArray()));
        }

        public String Trim3(string trimChars)
        {
            return new String(_string.Trim(trimChars.ToCharArray()));
        }

        public string TrimBStr()
        { 
            return _string.Trim(); 
        }

        public string TrimBStr2(string trimChars)
        {
            return _string.Trim(trimChars.ToCharArray());
        }

        public String TrimEnd(String trimChars)
        {
            return new String(_string.TrimEnd(trimChars.WrappedString.ToCharArray()));
        }

        public String TrimEnd2(string trimChars)
        { 
            return new String(_string.TrimEnd(trimChars.ToCharArray()));
        }

        public string TrimEndBStr(string trimChars)
        {
            return _string.TrimEnd(trimChars.ToCharArray());
        }

        public String TrimStart(String trimChars)
        {
            return new String(_string.TrimStart(trimChars.WrappedString.ToCharArray()));
        }

        public String TrimStart2(string trimChars)
        {
            return new String(_string.TrimStart(trimChars.ToCharArray()));
        }

        public string TrimStartBStr(string trimChars)
        { 
            return _string.TrimStart(trimChars.ToCharArray()); 
        }

        // Extensions
        public bool IsSurrogate()
        {
            return _string.IsSurrogate();
        }

    }
}
