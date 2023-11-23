// https://learn.microsoft.com/en-us/dotnet/api/system.text.utf8encoding?view=netframework-4.8.1

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
    [Description("Represents a UTF-8 encoding of Unicode characters.")]
    [Guid("B6C7F0B0-0D8F-4197-92DC-81F3BE308820")]
    [ProgId("DotNetLib.System.Text.UTF8Encoding")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IUTF8Encoding))]
    public class UTF8Encoding : IUTF8Encoding, Encoding, IWrappedObject
    {
        private GText.UTF8Encoding _UTF8Encoding;

        // Constructors

        internal UTF8Encoding(GText.UTF8Encoding utf8Encoding)
        {
            _UTF8Encoding = utf8Encoding;
        }

        public UTF8Encoding() 
        {
            _UTF8Encoding = new GText.UTF8Encoding();
        }

        public UTF8Encoding(bool encoderShouldEmitUTF8Identifier)
        {
            _UTF8Encoding = new GText.UTF8Encoding(encoderShouldEmitUTF8Identifier);
        }

        public UTF8Encoding(bool encoderShouldEmitUTF8Identifier, bool throwOnInvalidBytes)
        {
            _UTF8Encoding = new GText.UTF8Encoding(encoderShouldEmitUTF8Identifier, throwOnInvalidBytes);
        }

        // Properties

        internal GText.UTF8Encoding WrappedUTF8Encoding => _UTF8Encoding;
        public object WrappedObject => _UTF8Encoding;

        public string BodyName => _UTF8Encoding.BodyName;

        public int CodePage => _UTF8Encoding.CodePage;

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

        public string EncodingName => _UTF8Encoding.EncodingName;

        public string HeaderName => _UTF8Encoding.HeaderName;

        public bool IsBrowserDisplay =>  _UTF8Encoding.IsBrowserDisplay;

        public bool IsBrowserSave => _UTF8Encoding.IsBrowserSave;

        public bool IsMailNewsDisplay => _UTF8Encoding.IsMailNewsDisplay;

        public bool IsMailNewsSave => _UTF8Encoding.IsMailNewsSave;

        public bool IsReadOnly => _UTF8Encoding.IsReadOnly;

        public bool IsSingleByte => _UTF8Encoding.IsSingleByte;

        public string WebName => _UTF8Encoding.WebName;

        public int WindowsCodePage => _UTF8Encoding.WindowsCodePage;

        public GText.Encoding Encoding { get; }

        public object Clone()
        {
            return new UTF8Encoding((GText.UTF8Encoding)_UTF8Encoding.Clone());
        }

        public new bool Equals(object value)
        {
            return _UTF8Encoding.Equals(value.Unwrap());
        }

        public int GetByteCount(string s)
        {
            return _UTF8Encoding.GetByteCount(s);
        }

        public int GetByteCount(string s, int index, int count)
        {
            return _UTF8Encoding.GetByteCount(s.ToCharArray(), index, count);
        }

        public byte[] GetBytes(string s)
        {
            return _UTF8Encoding.GetBytes(s);
        }

        public virtual byte[] GetBytes(string s, int index, int count)
        {
            return _UTF8Encoding.GetBytes(s.ToCharArray(), index, count);
        }

        public int GetBytes(string s, int charIndex, int charCount, [In] ref byte[] bytes, int byteIndex)
        {
            return _UTF8Encoding.GetBytes(s, charIndex, charCount, bytes, byteIndex);
        }

        public int GetCharCount([In] ref byte[] bytes)
        {
            return _UTF8Encoding.GetCharCount(bytes);
        }

        public int GetCharCount([In] ref byte[] bytes, int index, int count)
        {
            return _UTF8Encoding.GetCharCount(bytes, index, count);
        }

        public Decoder GetDecoder()
        {
            return _UTF8Encoding.GetDecoder();
        }

        public Encoder GetEncoder()
        {
            return _UTF8Encoding.GetEncoder();
        }

        public override int GetHashCode()
        { 
            return _UTF8Encoding.GetHashCode(); 
        }

        public int GetMaxByteCount(int charCount)
        {
            return _UTF8Encoding.GetMaxByteCount(charCount);
        }

        public int GetMaxCharCount(int byteCount)
        {
            return _UTF8Encoding.GetMaxCharCount(byteCount);
        }

        public byte[] GetPreamble()
        {
            return _UTF8Encoding.GetPreamble();
        }

        public string GetString([In] ref byte[] bytes)
        {
            return _UTF8Encoding.GetString(bytes);
        }

        public string GetString([In] ref byte[] bytes, int index, int count)
        {
            return _UTF8Encoding.GetString(bytes, index, count);
        }

        public bool IsAlwaysNormalized()
        {
            return _UTF8Encoding.IsAlwaysNormalized();
        }

        public bool IsAlwaysNormalized(NormalizationForm form)
        {
            return _UTF8Encoding.IsAlwaysNormalized(form);
        }

        public new Type GetType()
        {
            return new Type(((GSystem.Object)this).GetType());
        }

        public new virtual string ToString()
        { 
            return _UTF8Encoding.ToString(); 
        }


    }
}
