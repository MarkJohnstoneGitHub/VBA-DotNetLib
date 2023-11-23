// https://learn.microsoft.com/en-us/dotnet/api/system.text.utf32encoding?view=netframework-4.8.1

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
    [Description("")]
    [Guid("F4CFC0B2-EB2B-457B-B7B5-5325B0F16B5B")]
    [ProgId("DotNetLib.System.Text.UTF32Encoding")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IUTF32Encoding))]
    public class UTF32Encoding : IUTF32Encoding, Encoding, IWrappedObject
    {
        private GText.UTF32Encoding _UTF32Encoding;

        // Constructors

        internal UTF32Encoding(GText.UTF32Encoding utf32Encoding)
        {
            _UTF32Encoding = utf32Encoding;
        }

        public UTF32Encoding()
        {
            _UTF32Encoding = new GText.UTF32Encoding();
        }

        public UTF32Encoding(bool bigEndian, bool byteOrderMark)
        {
            _UTF32Encoding = new GText.UTF32Encoding(bigEndian, byteOrderMark);
        }

        public UTF32Encoding(bool bigEndian, bool byteOrderMark, bool throwOnInvalidCharacters)
        {
            _UTF32Encoding = new GText.UTF32Encoding(bigEndian, byteOrderMark, throwOnInvalidCharacters);
        }


        // Properties

        public object WrappedObject => _UTF32Encoding;

        internal GText.UTF32Encoding WrappedUTF32Encoding => _UTF32Encoding;

        public string BodyName => _UTF32Encoding.BodyName;

        public int CodePage =>  _UTF32Encoding.CodePage;

        public DecoderFallback DecoderFallback 
        { 
            get => throw new NotImplementedException(); 
            set => throw new NotImplementedException(); 
        }
        public EncoderFallback EncoderFallback { 
            get => throw new NotImplementedException(); 
            set => throw new NotImplementedException(); 
        }

        public string EncodingName => _UTF32Encoding.EncodingName;

        public string HeaderName => _UTF32Encoding.HeaderName;

        public bool IsBrowserDisplay => _UTF32Encoding.IsBrowserDisplay;

        public bool IsBrowserSave => _UTF32Encoding.IsBrowserSave;

        public bool IsMailNewsDisplay => _UTF32Encoding.IsMailNewsDisplay;

        public bool IsMailNewsSave => _UTF32Encoding.IsMailNewsSave;

        public bool IsReadOnly => _UTF32Encoding.IsReadOnly;

        public bool IsSingleByte => _UTF32Encoding.IsSingleByte;

        public string WebName => _UTF32Encoding.WebName;
        public int WindowsCodePage => _UTF32Encoding.WindowsCodePage;

        // Methods

        public object Clone()
        {
            return new UTF32Encoding((GText.UTF32Encoding)_UTF32Encoding.Clone());
        }

        public new bool Equals(object value)
        {
            return _UTF32Encoding.Equals(value.Unwrap());
        }

        public int GetByteCount(string s)
        {
            return _UTF32Encoding.GetByteCount(s);
        }

        public int GetByteCount(string s, int index, int count)
        {
            return _UTF32Encoding.GetByteCount(s.ToCharArray(), index, count);
        }

        public byte[] GetBytes(string s)
        {
            return _UTF32Encoding.GetBytes(s);
        }

        public byte[] GetBytes(string s, int index, int count)
        {
            return _UTF32Encoding.GetBytes(s.ToCharArray(), index, count);
        }

        public int GetBytes(string s, int charIndex, int charCount, [In] ref byte[] bytes, int byteIndex)
        {
            return _UTF32Encoding.GetBytes(s, charIndex, charCount, bytes, byteIndex);
        }

        public int GetCharCount([In] ref byte[] bytes)
        {
            return _UTF32Encoding.GetCharCount(bytes);
        }

        public int GetCharCount([In] ref byte[] bytes, int index, int count)
        {
            return _UTF32Encoding.GetCharCount(bytes, index, count);
        }

        public Decoder GetDecoder()
        {
            return _UTF32Encoding.GetDecoder();
        }

        public Encoder GetEncoder()
        {
            return _UTF32Encoding.GetEncoder();
        }

        public override int GetHashCode()
        {
            return _UTF32Encoding.GetHashCode();
        }

        public int GetMaxByteCount(int charCount)
        {
            return _UTF32Encoding.GetMaxByteCount(charCount);
        }

        public int GetMaxCharCount(int byteCount)
        {
            return _UTF32Encoding.GetMaxCharCount(byteCount);
        }

        public byte[] GetPreamble()
        {
            return _UTF32Encoding.GetPreamble();
        }

        public string GetString([In] ref byte[] bytes)
        {
            return _UTF32Encoding.GetString(bytes);
        }

        public string GetString([In] ref byte[] bytes, int index, int count)
        {
            return _UTF32Encoding.GetString(bytes, index, count);
        }

        public bool IsAlwaysNormalized()
        {
            return _UTF32Encoding.IsAlwaysNormalized();
        }

        public bool IsAlwaysNormalized(NormalizationForm form)
        {
            return _UTF32Encoding.IsAlwaysNormalized(form);
        }

        public new Type GetType()
        {
            return new Type(((GSystem.Object)this).GetType());
        }

        public new virtual string ToString()
        {
            return _UTF32Encoding.ToString();
        }
    }
}
