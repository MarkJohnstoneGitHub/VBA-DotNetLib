// https://learn.microsoft.com/en-us/dotnet/api/system.text.encoding?view=netframework-4.8.1

using GText = global::System.Text;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Guid("5BF955D1-7A0C-4697-896B-2EF9A90C0F98")]
    [Description("Represents a character encoding.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IEncodingSingleton
    {
        // Properties
        Encoding ASCII 
        {
            [Description("Gets an encoding for the ASCII (7-bit) character set.")]
            get;
        }

        Encoding BigEndianUnicode 
        {
            [Description("Gets an encoding for the UTF-16 format that uses the big endian byte order.")]
            get;
        }

        Encoding Default 
        {
            [Description("Gets the default encoding for this .NET implementation.")]
            get;
        }


        Encoding Unicode
        {
            [Description("Gets an encoding for the UTF-16 format using the little endian byte order.")]
            get;
        }

        Encoding UTF32 
        {
            [Description("Gets an encoding for the UTF-32 format using the little endian byte order.")]
            get;
        }

        Encoding UTF7 
        {
            [Description("Gets an encoding for the UTF-7 format.")]
            get;
        }

        Encoding UTF8 
        {
            [Description("Gets an encoding for the UTF-8 format.")]
            get;
        }

        [Description("Converts an entire byte array from one encoding to another.")]
        byte[] Convert(Encoding srcEncoding, Encoding dstEncoding, [In] ref byte[] bytes);

        [Description("Converts a range of bytes in a byte array from one encoding to another.")]
        byte[] Convert(Encoding srcEncoding, Encoding dstEncoding, [In] ref byte[] bytes, int index, int count);

        [Description("Returns the encoding associated with the specified code page identifier.")]
        Encoding GetEncoding(int codepage);

        [Description("Returns the encoding associated with the specified code page name.")]
        Encoding GetEncoding(string name);

        [Description("Returns the encoding associated with the specified code page identifier. Parameters specify an error handler for characters that cannot be encoded and byte sequences that cannot be decoded.")]
        Encoding GetEncoding(int codepage, GText.EncoderFallback encoderFallback, GText.DecoderFallback decoderFallback);

        [Description("Returns the encoding associated with the specified code page name. Parameters specify an error handler for characters that cannot be encoded and byte sequences that cannot be decoded.")]
        Encoding GetEncoding(string name, GText.EncoderFallback encoderFallback, GText.DecoderFallback decoderFallback);

        [Description("Returns an array that contains all encodings.")]
        GText.EncodingInfo[] GetEncodings();

        [Description("Registers an encoding provider.")]
        void RegisterProvider(GText.EncodingProvider provider);

    }
}
