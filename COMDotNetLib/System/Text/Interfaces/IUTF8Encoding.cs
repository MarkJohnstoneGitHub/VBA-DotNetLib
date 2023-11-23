// https://learn.microsoft.com/en-us/dotnet/api/system.text.utf8encoding?view=netframework-4.8.1

using GText = global::System.Text;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Guid("0A5223F5-281E-4261-908F-AC7C09C4A2A0")]
    [Description("Represents a UTF-8 encoding of Unicode characters.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IUTF8Encoding
    {
        string BodyName
        {
            [Description("When overridden in a derived class, gets a name for the current encoding that can be used with mail agent body tags.\r\n\r\n(Inherited from Encoding)")]
            get;
        }

        int CodePage
        {
            [Description("When overridden in a derived class, gets the code page identifier of the current Encoding.\r\n\r\n(Inherited from Encoding)")]
            get;
        }

        GText.DecoderFallback DecoderFallback
        {
            [Description("Gets or sets the DecoderFallback object for the current Encoding object.\r\n\r\n(Inherited from Encoding)")]
            get;
            [Description("Gets or sets the DecoderFallback object for the current Encoding object.\r\n\r\n(Inherited from Encoding)")]
            set;
        }

        GText.EncoderFallback EncoderFallback
        {
            [Description("Gets or sets the EncoderFallback object for the current Encoding object.\r\n\r\n(Inherited from Encoding)")]
            get;
            [Description("Gets or sets the EncoderFallback object for the current Encoding object.\r\n\r\n(Inherited from Encoding)")]
            set;
        }

        string EncodingName
        {
            [Description("When overridden in a derived class, gets the human-readable description of the current encoding.\r\n\r\n(Inherited from Encoding)")]
            get;
        }

        string HeaderName
        {
            [Description("When overridden in a derived class, gets a name for the current encoding that can be used with mail agent header tags.\r\n\r\n(Inherited from Encoding)")]
            get;
        }

        bool IsBrowserDisplay
        {
            [Description("When overridden in a derived class, gets a value indicating whether the current encoding can be used by browser clients for displaying content.\r\n\r\n(Inherited from Encoding)")]
            get;
        }

        bool IsBrowserSave
        {
            [Description("When overridden in a derived class, gets a value indicating whether the current encoding can be used by browser clients for saving content.\r\n\r\n(Inherited from Encoding)")]
            get;
        }

        bool IsMailNewsDisplay
        {
            [Description("When overridden in a derived class, gets a value indicating whether the current encoding can be used by mail and news clients for displaying content.\r\n\r\n(Inherited from Encoding)")]
            get;
        }

        bool IsMailNewsSave
        {
            [Description("When overridden in a derived class, gets a value indicating whether the current encoding can be used by mail and news clients for saving content.\r\n\r\n(Inherited from Encoding)")]
            get;
        }

        bool IsReadOnly
        {
            [Description("When overridden in a derived class, gets a value indicating whether the current encoding is read-only.\r\n\r\n(Inherited from Encoding)")]
            get;
        }

        bool IsSingleByte
        {
            [Description("When overridden in a derived class, gets a value indicating whether the current encoding uses single-byte code points.\r\n\r\n(Inherited from Encoding)")]
            get;
        }

        string WebName
        {
            [Description("When overridden in a derived class, gets the name registered with the Internet Assigned Numbers Authority (IANA) for the current encoding.\r\n\r\n(Inherited from Encoding)")]
            get;
        }

        int WindowsCodePage
        {
            [Description("When overridden in a derived class, gets the Windows operating system code page that most closely corresponds to the current encoding.\r\n\r\n(Inherited from Encoding)")]
            get;
        }

        // Methods

        [Description("When overridden in a derived class, creates a shallow copy of the current Encoding object.\r\n\r\n(Inherited from Encoding)")]
        object Clone();

        [Description("Determines whether the specified object is equal to the current UTF8Encoding object.")]
        bool Equals(object value);

        [Description("Calculates the number of bytes produced by encoding the characters in the specified String.")]
        int GetByteCount(string s);

        [Description("When overridden in a derived class, calculates the number of bytes produced by encoding a set of characters from the specified string.\r\n\r\n(Inherited from Encoding)")] 
        int GetByteCount(string s, int index, int count);

        [Description("When overridden in a derived class, encodes all the characters in the specified string into a sequence of bytes.\r\n\r\n(Inherited from Encoding)")]
        byte[] GetBytes(string s);

        [Description("When overridden in a derived class, encodes into an array of bytes the number of characters specified by count in the specified string, starting from the specified index.\r\n\r\n(Inherited from Encoding)")]
        byte[] GetBytes(string s, int index, int count);

        [Description("Encodes a set of characters from the specified String into the specified byte array.")]
        int GetBytes(string s, int charIndex, int charCount, [In] ref byte[] bytes, int byteIndex);


        [Description("When overridden in a derived class, calculates the number of characters produced by decoding all the bytes in the specified byte array.\r\n\r\n(Inherited from Encoding)")]
        int GetCharCount([In] ref byte[] bytes);

        [Description("Calculates the number of characters produced by decoding a sequence of bytes from the specified byte array.")]
        int GetCharCount([In] ref byte[] bytes, int index, int count);

        [Description("Obtains a decoder that converts a UTF-8 encoded sequence of bytes into a sequence of Unicode characters.")]
        GText.Decoder GetDecoder();

        [Description("Obtains an encoder that converts a sequence of Unicode characters into a UTF-8 encoded sequence of bytes.")]
        GText.Encoder GetEncoder();

        [Description("Returns the hash code for the current instance.")]
        int GetHashCode();

        [Description("Calculates the maximum number of bytes produced by encoding the specified number of characters.")]
        int GetMaxByteCount(int charCount);

        [Description("Calculates the maximum number of characters produced by decoding the specified number of bytes.")]
        int GetMaxCharCount(int byteCount);

        [Description("Returns a Unicode byte order mark encoded in UTF-8 format, if the UTF8Encoding encoding object is configured to supply one.")]
        byte[] GetPreamble();

        [Description("When overridden in a derived class, decodes all the bytes in the specified byte array into a string.\r\n\r\n(Inherited from Encoding)")]
        string GetString([In] ref byte[] bytes);

        [Description("Decodes a range of bytes from a byte array into a string.")]
        string GetString([In] ref byte[] bytes, int index, int count);

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

        [Description("Gets a value indicating whether the current encoding is always normalized, using the default normalization form.\r\n\r\n(Inherited from Encoding)")]
        bool IsAlwaysNormalized();

        [Description("When overridden in a derived class, gets a value indicating whether the current encoding is always normalized, using the specified normalization form.\r\n\r\n(Inherited from Encoding)")]
        bool IsAlwaysNormalized(GText.NormalizationForm form);

        [Description("Returns a string that represents the current object.\r\n\r\n(Inherited from Object)")]
        string ToString();


    }
}
