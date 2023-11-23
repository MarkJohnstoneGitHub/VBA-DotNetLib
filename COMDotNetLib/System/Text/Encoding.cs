// https://learn.microsoft.com/en-us/dotnet/api/system.text.encoding?view=netframework-4.8.1

using GText = global::System.Text;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Guid("26290585-39FB-41D6-9C24-35F31985AA8A")]
    [Description("Represents a character encoding.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface Encoding
    {
        string BodyName 
        {
            [Description("When overridden in a derived class, gets a name for the current encoding that can be used with mail agent body tags.")]
            get;
        }

        int CodePage 
        {
            [Description("When overridden in a derived class, gets the code page identifier of the current Encoding.")]
            get;
        }

        GText.DecoderFallback DecoderFallback 
        {
            [Description("Gets or sets the DecoderFallback object for the current Encoding object.")]
            get;
            [Description("Gets or sets the DecoderFallback object for the current Encoding object.")]
            set;
        }

        GText.EncoderFallback EncoderFallback 
        {
            [Description("Gets or sets the EncoderFallback object for the current Encoding object.")]
            get;
            [Description("Gets or sets the EncoderFallback object for the current Encoding object.")]
            set;
        }

        string EncodingName 
        {
            [Description("When overridden in a derived class, gets the human-readable description of the current encoding.")]
            get;
        }

        string HeaderName 
        {
            [Description("When overridden in a derived class, gets a name for the current encoding that can be used with mail agent header tags.")]
            get;
        }

        bool IsBrowserDisplay 
        {
            [Description("When overridden in a derived class, gets a value indicating whether the current encoding can be used by browser clients for displaying content.")]
            get;
        }

        bool IsBrowserSave 
        {
            [Description("When overridden in a derived class, gets a value indicating whether the current encoding can be used by browser clients for saving content.")]
            get;
        }

        bool IsMailNewsDisplay 
        {
            [Description("When overridden in a derived class, gets a value indicating whether the current encoding can be used by mail and news clients for displaying content.")]
            get;
        }

        bool IsMailNewsSave 
        {
            [Description("When overridden in a derived class, gets a value indicating whether the current encoding can be used by mail and news clients for saving content.")]
            get;
        }

        bool IsReadOnly 
        {
            [Description("When overridden in a derived class, gets a value indicating whether the current encoding is read-only.")]
            get;
        }

        bool IsSingleByte 
        {
            [Description("When overridden in a derived class, gets a value indicating whether the current encoding uses single-byte code points.")]
            get;
        }

        string WebName 
        {
            [Description("When overridden in a derived class, gets the name registered with the Internet Assigned Numbers Authority (IANA) for the current encoding.")]
            get;
        }

        int WindowsCodePage 
        {
            [Description("When overridden in a derived class, gets the Windows operating system code page that most closely corresponds to the current encoding.")]
            get;
        }

        // Methods

        [Description("When overridden in a derived class, creates a shallow copy of the current Encoding object.")]
        object Clone();

        [Description("Determines whether the specified Object is equal to the current instance.")]
        bool Equals(object value);

        [Description("When overridden in a derived class, calculates the number of bytes produced by encoding the characters in the specified string")]
        int GetByteCount(string s);

        [Description("When overridden in a derived class, calculates the number of bytes produced by encoding a set of characters from the specified string.")]
        int GetByteCount(string s, int index, int count);

        [Description("When overridden in a derived class, encodes all the characters in the specified string into a sequence of bytes.")]
        byte[] GetBytes(string s);

        [Description("When overridden in a derived class, encodes into an array of bytes the number of characters specified by count in the specified string, starting from the specified index.")]
        byte[] GetBytes(string s, int index, int count);

        [Description("When overridden in a derived class, encodes a set of characters from the specified string into the specified byte array.")]
        int GetBytes(string s, int charIndex, int charCount, [In] ref byte[] bytes, int byteIndex);

        [Description("When overridden in a derived class, calculates the number of characters produced by decoding all the bytes in the specified byte array.")]
        int GetCharCount([In] ref byte[] bytes);

        [Description("When overridden in a derived class, calculates the number of characters produced by decoding a sequence of bytes from the specified byte array.")]
        int GetCharCount([In] ref byte[] bytes, int index, int count);

        //[Description("When overridden in a derived class, decodes all the bytes in the specified byte array into a set of characters.")]
        //string GetChars(byte[] bytes);

        //[Description("When overridden in a derived class, decodes a sequence of bytes from the specified byte array into a set of characters.")]
        //string GetChars(byte[] bytes, int index, int count);

        [Description("When overridden in a derived class, obtains a decoder that converts an encoded sequence of bytes into a sequence of characters.")]
        GText.Decoder GetDecoder();

        [Description("When overridden in a derived class, obtains an encoder that converts a sequence of Unicode characters into an encoded sequence of bytes.")]
        GText.Encoder GetEncoder();

        [Description("Returns the hash code for the current instance.")]
        int GetHashCode();

        [Description("When overridden in a derived class, calculates the maximum number of bytes produced by encoding the specified number of characters.")]
        int GetMaxByteCount(int charCount);

        [Description("When overridden in a derived class, calculates the maximum number of characters produced by decoding the specified number of bytes.")]
        int GetMaxCharCount(int byteCount);

        [Description("When overridden in a derived class, returns a sequence of bytes that specifies the encoding used.")]
        byte[] GetPreamble();

        [Description("When overridden in a derived class, decodes all the bytes in the specified byte array into a string.")]
        string GetString([In] ref byte[] bytes);

        [Description("When overridden in a derived class, decodes a sequence of bytes from the specified byte array into a string.")]
        string GetString([In] ref byte[] bytes, int index, int count);

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

        [Description("Gets a value indicating whether the current encoding is always normalized, using the default normalization form.")]
        bool IsAlwaysNormalized();

        [Description("When overridden in a derived class, gets a value indicating whether the current encoding is always normalized, using the specified normalization form.")]
        bool IsAlwaysNormalized(GText.NormalizationForm form);

        [Description("Returns a string that represents the current object.\r\n\r\n(Inherited from Object)")]
        string ToString();



        //[Description("")]
    }
}
