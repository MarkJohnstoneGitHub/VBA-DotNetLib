// https://learn.microsoft.com/en-us/dotnet/api/system.text.encoding?view=netframework-4.8.1

using DotNetLib.Extensions;
using GText = global::System.Text;
using System.Runtime.InteropServices;
using System.Security;
using System.ComponentModel;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Description("Represents a character encoding.")]
    [Guid("469A677D-25F4-4105-877E-0E7D9C5BF529")]
    [ProgId("DotNetLib.System.Text.EncodingSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IEncodingSingleton))]

    public class EncodingSingleton : IEncodingSingleton
    {
        private static volatile Encoding asciiEncoding;

        private static volatile Encoding defaultEncoding;

        private static volatile Encoding unicodeEncoding;

        private static volatile Encoding bigEndianUnicode;

        private static volatile Encoding utf7Encoding;

        private static volatile Encoding utf8Encoding;

        private static volatile Encoding utf32Encoding;

        public EncodingSingleton() { }

        public Encoding ASCII
        {
            get
            {
                if (asciiEncoding == null)
                {
                    asciiEncoding = new ASCIIEncoding();
                }

                return asciiEncoding;
            }
        }

        public Encoding BigEndianUnicode
        {
            get
            {
                if (bigEndianUnicode == null)
                {
                    bigEndianUnicode = new UnicodeEncoding(bigEndian: true, byteOrderMark: true);
                }

                return bigEndianUnicode;
            }
        }

        public Encoding Default
        {
            [SecuritySafeCritical]
            get
            {
                if (defaultEncoding == null)
                {
                    defaultEncoding = GText.Encoding.Default.Wrap();
                }

                return defaultEncoding;
            }
        }

        public Encoding Unicode
        {
            get
            {
                if (unicodeEncoding == null)
                {
                    unicodeEncoding = new UnicodeEncoding(bigEndian: false, byteOrderMark: true);
                }

                return unicodeEncoding;
            }
        }

        public Encoding UTF32
        {
            get
            {
                if (utf32Encoding == null)
                {
                    utf32Encoding = new UTF32Encoding(bigEndian: false, byteOrderMark: true);
                }

                return utf32Encoding;
            }
        }

        public Encoding UTF7
        {
            get
            {
                if (utf7Encoding == null)
                {
                    utf7Encoding = new UTF7Encoding();
                }

                return utf7Encoding;
            }
        }

        public Encoding UTF8
        {
            get
            {
                if (utf8Encoding == null)
                {
                    utf8Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);
                }

                return utf8Encoding;
            }
        }

        // Methods

        public byte[] Convert(Encoding srcEncoding, Encoding dstEncoding, [In] ref byte[] bytes)
        {
            return GText.Encoding.Convert(srcEncoding.UnWrapEncoding(), dstEncoding.UnWrapEncoding(), bytes);
        }

        public byte[] Convert(Encoding srcEncoding, Encoding dstEncoding, [In] ref byte[] bytes, int index, int count)
        {
            return GText.Encoding.Convert(srcEncoding.UnWrapEncoding(), dstEncoding.UnWrapEncoding(), bytes, index, count);
        }

        public Encoding GetEncoding(int codepage)
        {
            return GText.Encoding.GetEncoding(codepage).Wrap();
        }

        public Encoding GetEncoding(string name)
        {
            return GText.Encoding.GetEncoding(name).Wrap();
        }

        public Encoding GetEncoding(int codepage, GText.EncoderFallback encoderFallback, GText.DecoderFallback decoderFallback)
        {
            return GText.Encoding.GetEncoding(codepage, encoderFallback, decoderFallback).Wrap();
        }

        public Encoding GetEncoding(string name, GText.EncoderFallback encoderFallback, GText.DecoderFallback decoderFallback)
        {
            return GText.Encoding.GetEncoding(name, encoderFallback, decoderFallback).Wrap();
        }

        public GText.EncodingInfo[] GetEncodings()
        { 
            return GText.Encoding.GetEncodings();
        }

        public void RegisterProvider(GText.EncodingProvider provider)
        {
            GText.Encoding.RegisterProvider(provider);
        }

    }
}
