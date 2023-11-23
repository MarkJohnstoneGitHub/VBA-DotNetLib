// https://learn.microsoft.com/en-us/dotnet/api/system.text.unicodeencoding?view=netframework-4.8.1

using GSystem = global::System;
using GText = global::System.Text;
using System;
using System.Text;
using System.Runtime.InteropServices;
using DotNetLib.Extensions;
using System.ComponentModel;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Description("Represents a UTF-16 encoding of Unicode characters.")]
    [Guid("282B1112-0566-4B66-9597-23F86CEF8069")]
    [ProgId("DotNetLib.System.Text.UnicodeEncoding")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IUnicodeEncoding))]
    public class UnicodeEncoding : IUnicodeEncoding, Encoding, IWrappedObject
    {
        private GText.UnicodeEncoding _UnicodeEncoding;

        // Constructors

        internal UnicodeEncoding(GText.UnicodeEncoding unicodeEncoding)
        {
            _UnicodeEncoding = unicodeEncoding;
        }

        public UnicodeEncoding() 
        {
            _UnicodeEncoding = new GText.UnicodeEncoding();
        }

        public UnicodeEncoding(bool bigEndian, bool byteOrderMark)
        {
            _UnicodeEncoding = new GText.UnicodeEncoding(bigEndian, byteOrderMark);
        }

        public UnicodeEncoding(bool bigEndian, bool byteOrderMark, bool throwOnInvalidBytes)
        {
            _UnicodeEncoding = new GText.UnicodeEncoding(bigEndian, byteOrderMark, throwOnInvalidBytes);
        }

        // Properties
        internal GText.UnicodeEncoding WrappedUnicodeEncoding => _UnicodeEncoding;
        public object WrappedObject => _UnicodeEncoding;

        public string BodyName => _UnicodeEncoding.BodyName;

        public int CodePage => _UnicodeEncoding.CodePage;

        public DecoderFallback DecoderFallback 
        { 
            get => throw new NotImplementedException(); 
            set => throw new NotImplementedException(); 
        }

        public EncoderFallback EncoderFallback 
        { 
            get => throw new NotImplementedException();
            set => throw new NotImplementedException(); 
        }

        public string EncodingName => _UnicodeEncoding.EncodingName;

        public string HeaderName => _UnicodeEncoding.HeaderName;

        public bool IsBrowserDisplay => _UnicodeEncoding.IsBrowserDisplay;

        public bool IsBrowserSave => _UnicodeEncoding.IsBrowserSave;

        public bool IsMailNewsDisplay => _UnicodeEncoding.IsMailNewsDisplay;

        public bool IsMailNewsSave => _UnicodeEncoding.IsMailNewsSave;

        public bool IsReadOnly => _UnicodeEncoding.IsReadOnly;

        public bool IsSingleByte => _UnicodeEncoding.IsSingleByte;

        public string WebName => _UnicodeEncoding.WebName;

        public int WindowsCodePage =>  _UnicodeEncoding.WindowsCodePage;

        // Methods

        public object Clone()
        {
            return new UnicodeEncoding((GText.UnicodeEncoding)_UnicodeEncoding.Clone());
        }

        public new bool Equals(object value)
        {
            return _UnicodeEncoding.Equals(value.Unwrap());
        }

        public int GetByteCount(string s)
        {
            return _UnicodeEncoding.GetByteCount(s);
        }

        public int GetByteCount(string s, int index, int count)
        {
            return _UnicodeEncoding.GetByteCount(s.ToCharArray(), index, count);
        }

        public byte[] GetBytes(string s)
        {
            return _UnicodeEncoding.GetBytes(s);
        }

        public virtual byte[] GetBytes(string s, int index, int count)
        {
            return _UnicodeEncoding.GetBytes(s.ToCharArray(), index, count);
        }

        public int GetBytes(string s, int charIndex, int charCount, [In] ref byte[] bytes, int byteIndex)
        {
            return _UnicodeEncoding.GetBytes(s, charIndex, charCount, bytes, byteIndex);
        }

        public int GetCharCount([In] ref byte[] bytes)
        {
            return _UnicodeEncoding.GetCharCount(bytes);
        }

        public int GetCharCount([In] ref byte[] bytes, int index, int count)
        {
            return _UnicodeEncoding.GetCharCount(bytes, index, count);
        }

        public Decoder GetDecoder()
        {
            return _UnicodeEncoding.GetDecoder();
        }

        public Encoder GetEncoder()
        {
            return _UnicodeEncoding.GetEncoder();
        }

        public override int GetHashCode()
        {
            return _UnicodeEncoding.GetHashCode();
        }

        public int GetMaxByteCount(int charCount)
        {
            return _UnicodeEncoding.GetMaxByteCount(charCount);
        }

        public int GetMaxCharCount(int byteCount)
        {
            return _UnicodeEncoding.GetMaxCharCount(byteCount);
        }

        public byte[] GetPreamble()
        {
            return _UnicodeEncoding.GetPreamble();
        }

        public string GetString([In] ref byte[] bytes)
        {
            return _UnicodeEncoding.GetString(bytes);
        }

        public string GetString([In] ref byte[] bytes, int index, int count)
        {
            return _UnicodeEncoding.GetString(bytes, index, count);
        }

        public bool IsAlwaysNormalized()
        {
            return _UnicodeEncoding.IsAlwaysNormalized();
        }

        public bool IsAlwaysNormalized(NormalizationForm form)
        {
            return _UnicodeEncoding.IsAlwaysNormalized(form);
        }

        public new Type GetType()
        {
            return new Type(((GSystem.Object)this).GetType());
        }

        public new virtual string ToString()
        {
            return _UnicodeEncoding.ToString();
        }

    }
}
