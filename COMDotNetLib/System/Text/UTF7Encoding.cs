// https://learn.microsoft.com/en-us/dotnet/api/system.text.utf7encoding?view=netframework-4.8.1

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
    [Description("Represents a UTF-7 encoding of Unicode characters.")]
    [Guid("DBDAE439-E236-446D-A55C-3F7A50D657D0")]
    [ProgId("DotNetLib.System.Text.UTF7Encoding")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IUTF7Encoding))]
    public class UTF7Encoding : IUTF7Encoding, Encoding, IWrappedObject
    {
        private GText.UTF7Encoding _UTF7Encoding;

        // Constructors
        internal UTF7Encoding(GText.UTF7Encoding utf7Encoding)
        {
            _UTF7Encoding = utf7Encoding;
        }

        public UTF7Encoding() 
        {
            _UTF7Encoding = new GText.UTF7Encoding();
        }

        public UTF7Encoding(bool allowOptionals)
        {
            _UTF7Encoding = new GText.UTF7Encoding(allowOptionals);
        }

        internal GText.UTF7Encoding WrappedUTF7Encoding => _UTF7Encoding;
        public object WrappedObject => _UTF7Encoding;

        public string BodyName => throw new NotImplementedException();

        public int CodePage => throw new NotImplementedException();

        public DecoderFallback DecoderFallback { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public EncoderFallback EncoderFallback { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public string EncodingName => _UTF7Encoding.EncodingName;
        public string HeaderName => _UTF7Encoding.HeaderName;

        public bool IsBrowserDisplay => _UTF7Encoding.IsBrowserDisplay;

        public bool IsBrowserSave => _UTF7Encoding.IsBrowserSave;

        public bool IsMailNewsDisplay => _UTF7Encoding.IsMailNewsDisplay;

        public bool IsMailNewsSave => _UTF7Encoding.IsMailNewsSave;

        public bool IsReadOnly => _UTF7Encoding.IsReadOnly;

        public bool IsSingleByte => _UTF7Encoding.IsSingleByte;

        public string WebName => _UTF7Encoding.WebName;

        public int WindowsCodePage => _UTF7Encoding.WindowsCodePage;

        // Methods

        public object Clone()
        {
            return new UTF7Encoding((GText.UTF7Encoding)_UTF7Encoding.Clone());
        }

        public new bool Equals(object value)
        {
            return _UTF7Encoding.Equals(value.Unwrap());
        }

        public int GetByteCount(string s)
        {
            return _UTF7Encoding.GetByteCount(s);
        }

        public int GetByteCount(string s, int index, int count)
        {
            return _UTF7Encoding.GetByteCount(s.ToCharArray(), index, count);
        }

        public byte[] GetBytes(string s)
        {
            return _UTF7Encoding.GetBytes(s);
        }

        public byte[] GetBytes(string s, int index, int count)
        {
            return _UTF7Encoding.GetBytes(s.ToCharArray(), index, count);
        }

        public int GetBytes(string s, int charIndex, int charCount, [In] ref byte[] bytes, int byteIndex)
        {
            return _UTF7Encoding.GetBytes(s, charIndex, charCount, bytes, byteIndex); 
        }

        public int GetCharCount([In] ref byte[] bytes)
        {
           return _UTF7Encoding.GetCharCount(bytes);
        }

        public int GetCharCount([In] ref byte[] bytes, int index, int count)
        {
            return _UTF7Encoding.GetCharCount(bytes, index, count);
        }

        public Decoder GetDecoder()
        {
            return _UTF7Encoding.GetDecoder();
        }

        public Encoder GetEncoder()
        {
            return _UTF7Encoding.GetEncoder();
        }

        public override int GetHashCode()
        {
            return _UTF7Encoding.GetHashCode();
        }

        public int GetMaxByteCount(int charCount)
        {
            return _UTF7Encoding.GetMaxByteCount(charCount);
        }

        public int GetMaxCharCount(int byteCount)
        {
            return _UTF7Encoding.GetMaxCharCount(byteCount);
        }

        public byte[] GetPreamble()
        {
            return _UTF7Encoding.GetPreamble();
        }

        public string GetString([In] ref byte[] bytes)
        {
            return _UTF7Encoding.GetString(bytes);
        }

        public string GetString([In] ref byte[] bytes, int index, int count)
        {
            return _UTF7Encoding.GetString(bytes, index, count);
        }

        public bool IsAlwaysNormalized()
        {
            return _UTF7Encoding.IsAlwaysNormalized();
        }

        public bool IsAlwaysNormalized(NormalizationForm form)
        {
            return _UTF7Encoding.IsAlwaysNormalized(form);
        }

        public new Type GetType()
        {
            return new Type(((GSystem.Object)this).GetType());
        }

        public new virtual string ToString()
        {
            return _UTF7Encoding.ToString();
        }
    }
}
