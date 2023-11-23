//  https://learn.microsoft.com/en-us/dotnet/api/system.text.encoding?view=netframework-4.8.1

using GText = global::System.Text;
using System.Runtime.InteropServices;
using UTF8Encoding = DotNetLib.System.Text.UTF8Encoding;
using Encoding = DotNetLib.System.Text.Encoding;
using UTF7Encoding = DotNetLib.System.Text.UTF7Encoding;
using UTF32Encoding = DotNetLib.System.Text.UTF32Encoding;
using UnicodeEncoding = DotNetLib.System.Text.UnicodeEncoding;
using ASCIIEncoding = DotNetLib.System.Text.ASCIIEncoding;

namespace DotNetLib.Extensions
{
    [ComVisible(false)]
    public static class EncodingExtension
    {
        public static Encoding Wrap(this GText.Encoding encoding)
        {
            if (encoding == null)
            {
                return null;
            }

            if (encoding is GText.UTF8Encoding utf8Encoding)
            {
                return new UTF8Encoding(utf8Encoding);
            }

            if (encoding is GText.UTF7Encoding utf7Encoding)
            {
                return new UTF7Encoding(utf7Encoding);
            }

            if (encoding is GText.UTF32Encoding utf32Encoding)
            {
                return new UTF32Encoding(utf32Encoding);
            }

            if (encoding is GText.UnicodeEncoding unicodeEncoding)
            {
                return new UnicodeEncoding(unicodeEncoding);
            }

            if (encoding is GText.ASCIIEncoding asciiEncoding)
            {
                return new ASCIIEncoding(asciiEncoding);
            }

            return null; //If encoding COM wrapper not implemented return null
        }

        public static GText.Encoding UnWrapEncoding(this Encoding encoding)
        {
            if (encoding == null)
            {
                return null;
            }

            if (encoding is UTF8Encoding utf8Encoding)
            {
                return utf8Encoding.WrappedUTF8Encoding;
            }

            if (encoding is UTF7Encoding utf7Encoding)
            {
                return utf7Encoding.WrappedUTF7Encoding;
            }

            if (encoding is UTF32Encoding utf32Encoding)
            {
                return utf32Encoding.WrappedUTF32Encoding;
            }

            if (encoding is UnicodeEncoding unicodeEncoding)
            {
                return unicodeEncoding.WrappedUnicodeEncoding;
            }

            if (encoding is ASCIIEncoding asciiEncoding)
            {
                return asciiEncoding.WrappedASCIIEncoding;
            }

            return null; //If encoding COM wrapper not implemented return null
        }
    }
}
