﻿// https://learn.microsoft.com/en-us/dotnet/api/system.char?view=netframework-4.8.1

using GSystem = global::System;
using GGlobalization = global::System.Globalization;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Description("Represents a character as a UTF-16 code unit.")]
    [Guid("7B001A76-5C0C-4F21-B3E4-B6FDC8632BA9")]
    [ProgId("DotNetLib.System.CharSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICharSingleton))]
    public class CharSingleton : ICharSingleton
    {
        public CharSingleton() { }

        // Fields

        //public void MaxValue([In][Out] ref char maxValue)
        //{
        //    maxValue = char.MaxValue; 
        //}
        //public void MinValue([In][Out] ref char minValue)
        //{
        //    minValue = char.MaxValue;
        //}

        public string ConvertFromUtf32(int utf32)
        {  
            return GSystem.Char.ConvertFromUtf32(utf32); 
        }

        public int ConvertToUtf32(string s, int index)
        {
            return GSystem.Char.ConvertToUtf32(s, index);
        }

        //public int ConvertToUtf32([In] ref char highSurrogate,[In] ref char lowSurrogate)
        //{
        //    return char.ConvertToUtf32(highSurrogate, lowSurrogate);
        //}


        public double GetNumericValue(string s, int index)
        {
            return GSystem.Char.GetNumericValue(s, index);
        }

        public GGlobalization.UnicodeCategory GetUnicodeCategory(string s, int index)
        {
            return GSystem.Char.GetUnicodeCategory(s, index);
        }

        public bool IsControl(string s, int index)
        {
            return GSystem.Char.IsControl(s, index);
        }

        public bool IsDigit(string s, int index)
        {
            return GSystem.Char.IsDigit(s, index);
        }

        public bool IsHighSurrogate(int charCode)
        {
            return GSystem.Char.IsHighSurrogate((char)charCode);
        }

        public bool IsHighSurrogate2(string s, int index)
        {
            return GSystem.Char.IsHighSurrogate(s, index);
        }

        public bool IsLetter(string s, int index)
        {
            return GSystem.Char.IsLetter(s, index);
        }

        public bool IsLetterOrDigit(string s, int index)
        {
            return GSystem.Char.IsLetterOrDigit(s, index);
        }

        public bool IsLower(string s, int index)
        {
            return GSystem.Char.IsLower(s, index);
        }

        public bool IsLowSurrogate(int charCode)
        {
            return GSystem.Char.IsLowSurrogate((char)charCode);
        }

        public bool IsLowSurrogate2(string s, int index)
        {
            return GSystem.Char.IsLowSurrogate(s, index);
        }

        public bool IsNumber(string s, int index)
        {
            return GSystem.Char.IsNumber(s, index);
        }

        public bool IsPunctuation(string s, int index)
        {
            return GSystem.Char.IsPunctuation(s, index);
        }

        public bool IsSeparator(string s, int index)
        {
            return GSystem.Char.IsSeparator(s, index);
        }
        public bool IsSurrogate(string s, int index)
        {
            return GSystem.Char.IsSurrogate(s, index);
        }

        // https://stackoverflow.com/questions/289792/int-to-char-in-charCode-sharp
        public bool IsSurrogate2(int charCode)
        {
            return GSystem.Char.IsSurrogate((char)charCode);
        }

        public bool IsSurrogatePair(string s, int index)
        {
            return GSystem.Char.IsSurrogatePair(s, index);
        }

        public bool IsSurrogatePair2(int highSurrogate, int lowSurrogate)
        {
            return GSystem.Char.IsSurrogatePair((char)highSurrogate, (char)lowSurrogate);
        }

        public bool IsSymbol(string s, int index)
        {
            return GSystem.Char.IsSymbol(s, index);
        }

        public bool IsUpper(string s, int index)
        {
            return GSystem.Char.IsUpper(s, index);
        }

        public bool IsWhiteSpace(string s, int index)
        {
            return GSystem.Char.IsWhiteSpace(s, index);
        }

        //public string ToString2([In] ref char charCode)
        //{
        //    return char.ToString2(charCode);
        //}

        //Operators
        //public bool Equality([In] ref char left, [In] ref char right)
        //{
        //    return (left == right);
        //}

        //public bool Inequality([In] ref char left, [In] ref char right)
        //{ 
        //    return (left != right);
        //}

        //public bool LessThan([In] ref char left, [In] ref char right)
        //{
        //    return (left < right);
        //}

        //public bool GreaterThan([In] ref char left, [In] ref char right)
        //{
        //    return (left > right);
        //}

    }
}
