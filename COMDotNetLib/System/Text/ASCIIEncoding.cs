// https://learn.microsoft.com/en-us/dotnet/api/system.text.asciiencoding?view=netframework-4.8.1

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
    [Description("Represents an ASCII character encoding of Unicode characters.")]
    [Guid("2753D356-3934-418E-BEF4-2630833DDD53")]
    [ProgId("DotNetLib.System.Text.ASCIIEncoding")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IASCIIEncoding))]
    public class ASCIIEncoding : IASCIIEncoding, Encoding, IWrappedObject
    {
        private GText.ASCIIEncoding _ASCIIEncoding;

        // Constructors
        internal ASCIIEncoding(GText.ASCIIEncoding asciiEncoding)
        {
            _ASCIIEncoding = asciiEncoding;
        }

        public ASCIIEncoding()
        {
            _ASCIIEncoding = new GText.ASCIIEncoding();
        }

        // Properties
        internal GText.ASCIIEncoding WrappedASCIIEncoding => _ASCIIEncoding;
        public object WrappedObject => _ASCIIEncoding;

        public string BodyName => _ASCIIEncoding.BodyName;

        public int CodePage => _ASCIIEncoding.CodePage;

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

        public string EncodingName => _ASCIIEncoding.EncodingName;

        public string HeaderName => _ASCIIEncoding.HeaderName;

        public bool IsBrowserDisplay => _ASCIIEncoding.IsBrowserDisplay;

        public bool IsBrowserSave => _ASCIIEncoding.IsBrowserSave;

        public bool IsMailNewsDisplay => _ASCIIEncoding.IsMailNewsDisplay;

        public bool IsMailNewsSave => _ASCIIEncoding.IsMailNewsSave;

        public bool IsReadOnly => _ASCIIEncoding.IsReadOnly;

        public bool IsSingleByte => _ASCIIEncoding.IsSingleByte;

        public string WebName => _ASCIIEncoding.WebName;

        public int WindowsCodePage => _ASCIIEncoding.WindowsCodePage;

        public object Clone()
        {
            return new ASCIIEncoding((GText.ASCIIEncoding)_ASCIIEncoding.Clone());
        }

        public new bool Equals(object value)
        {
            return _ASCIIEncoding.Equals(value.Unwrap());
        }

        public int GetByteCount(string s)
        {
            return _ASCIIEncoding.GetByteCount(s);
        }

        public int GetByteCount(string s, int index, int count)
        {
            return _ASCIIEncoding.GetByteCount(s.ToCharArray(), index, count);
        }

        public byte[] GetBytes(string s)
        {
            return _ASCIIEncoding.GetBytes(s);
        }

        public virtual byte[] GetBytes(string s, int index, int count)
        {
            return _ASCIIEncoding.GetBytes(s.ToCharArray(), index, count);
        }

        public int GetBytes(string s, int charIndex, int charCount, [In] ref byte[] bytes, int byteIndex)
        {
            return _ASCIIEncoding.GetBytes(s,charIndex, charCount, bytes, byteIndex);
        }

        public int GetCharCount([In] ref byte[] bytes)
        {
            return _ASCIIEncoding.GetCharCount(bytes);
        }

        public int GetCharCount([In] ref byte[] bytes, int index, int count)
        {
            return _ASCIIEncoding.GetCharCount(bytes, index, count);
        }

        public Decoder GetDecoder()
        {
            return _ASCIIEncoding.GetDecoder();
        }

        public Encoder GetEncoder()
        {
            return _ASCIIEncoding.GetEncoder();
        }

        public override int GetHashCode()
        {
            return _ASCIIEncoding.GetHashCode();
        }

        public int GetMaxByteCount(int charCount)
        {
           return _ASCIIEncoding.GetMaxByteCount(charCount);
        }

        public int GetMaxCharCount(int byteCount)
        {
            return _ASCIIEncoding.GetMaxCharCount(byteCount);
        }

        public byte[] GetPreamble()
        {
            return _ASCIIEncoding.GetPreamble();
        }

        public string GetString([In] ref byte[] bytes)
        {
            return _ASCIIEncoding.GetString(bytes);
        }

        public string GetString([In] ref byte[] bytes, int index, int count)
        {
            return _ASCIIEncoding.GetString(bytes, index, count);
        }

        public new Type GetType()
        {
            return new Type(((GSystem.Object)this).GetType());
        }

        public bool IsAlwaysNormalized()
        {
            return _ASCIIEncoding.IsAlwaysNormalized();
        }

        public bool IsAlwaysNormalized(NormalizationForm form)
        {
            return _ASCIIEncoding.IsAlwaysNormalized(form);
        }

        public new virtual string ToString()
        {
            return _ASCIIEncoding.ToString();
        }

    }
}
