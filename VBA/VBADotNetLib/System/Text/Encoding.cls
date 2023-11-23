VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Encoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a character encoding."
'@Folder "VBADotNetLib.System.Text"
'@ModuleDescription("Represents a character encoding.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 23, 2023
'@Version v1.0 November 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.encoding?view=netframework-4.8.1

Option Explicit

Private Type TEncodingSingleton
   EncodingSingleton As DotNetLib.EncodingSingleton
End Type

Private this As TEncodingSingleton

Private Sub Class_Initialize()
   Set this.EncodingSingleton = New DotNetLib.EncodingSingleton
End Sub

Private Sub Class_Terminate()
   Set this.EncodingSingleton = Nothing
End Sub

'@Description("Gets an encoding for the ASCII (7-bit) character set.")
Public Property Get ASCII() As DotNetLib.Encoding
Attribute ASCII.VB_Description = "Gets an encoding for the ASCII (7-bit) character set."
   Set ASCII = this.EncodingSingleton.ASCII()
End Property

'@Description("Gets an encoding for the UTF-16 format that uses the big endian byte order.")
Public Property Get BigEndianUnicode() As DotNetLib.Encoding
Attribute BigEndianUnicode.VB_Description = "Gets an encoding for the UTF-16 format that uses the big endian byte order."
   Set BigEndianUnicode = this.EncodingSingleton.BigEndianUnicode()
End Property

'@Description("Gets the default encoding for this .NET implementation.")
Public Property Get Default() As DotNetLib.Encoding
Attribute Default.VB_Description = "Gets the default encoding for this .NET implementation."
   Set Default = this.EncodingSingleton.Default()
End Property

'@Description("Gets an encoding for the UTF-16 format using the little endian byte order.")
Public Property Get Unicode() As DotNetLib.Encoding
Attribute Unicode.VB_Description = "Gets an encoding for the UTF-16 format using the little endian byte order."
   Set Unicode = this.EncodingSingleton.Unicode()
End Property

'@Description("Gets an encoding for the UTF-32 format using the little endian byte order.")
Public Property Get Utf32() As DotNetLib.Encoding
Attribute Utf32.VB_Description = "Gets an encoding for the UTF-32 format using the little endian byte order."
   Set Utf32 = this.EncodingSingleton.Utf32()
End Property

'@Description("Gets an encoding for the UTF-7 format.")
Public Property Get UTF7() As DotNetLib.Encoding
Attribute UTF7.VB_Description = "Gets an encoding for the UTF-7 format."
   Set UTF7 = this.EncodingSingleton.UTF7()
End Property

'@Description("Gets an encoding for the UTF-8 format.")
Public Property Get UTF8() As DotNetLib.Encoding
Attribute UTF8.VB_Description = "Gets an encoding for the UTF-8 format."
   Set UTF8 = this.EncodingSingleton.UTF8()
End Property

'@Description("Converts an entire byte array from one encoding to another.")
Public Function Convert(ByVal srcEncoding As DotNetLib.Encoding, ByVal dstEncoding As DotNetLib.Encoding, ByRef bytes() As Byte) As Byte()
Attribute Convert.VB_Description = "Converts an entire byte array from one encoding to another."
   Convert = this.EncodingSingleton.Convert(srcEncoding, dstEncoding, bytes)
End Function

'@Description("Converts a range of bytes in a byte array from one encoding to another.")
Public Function Convert2(ByVal srcEncoding As DotNetLib.Encoding, ByVal dstEncoding As DotNetLib.Encoding, ByRef bytes() As Byte, ByVal index As Long, ByVal count As Long) As Byte()
Attribute Convert2.VB_Description = "Converts a range of bytes in a byte array from one encoding to another."
   Convert2 = this.EncodingSingleton.Convert_2(srcEncoding, dstEncoding, bytes, index, count)
End Function

'@Description("Returns the encoding associated with the specified code page identifier.")
Public Function GetEncoding(ByVal pCodePage As Long) As DotNetLib.Encoding
Attribute GetEncoding.VB_Description = "Returns the encoding associated with the specified code page identifier."
   Set GetEncoding = this.EncodingSingleton.GetEncoding(pCodePage)
End Function

'@Description("Returns the encoding associated with the specified code page name.")
Public Function GetEncoding2(ByVal name As String) As DotNetLib.Encoding
Attribute GetEncoding2.VB_Description = "Returns the encoding associated with the specified code page name."
   Set GetEncoding2 = this.EncodingSingleton.GetEncoding_2(name)
End Function

'@Description("Returns the encoding associated with the specified code page identifier. Parameters specify an error handler for characters that cannot be encoded and byte sequences that cannot be decoded.")
Public Function GetEncoding3(ByVal pCodePage As Long, ByVal EncoderFallback As Object, ByVal DecoderFallback As Object) As DotNetLib.Encoding
Attribute GetEncoding3.VB_Description = "Returns the encoding associated with the specified code page identifier. Parameters specify an error handler for characters that cannot be encoded and byte sequences that cannot be decoded."
   Set GetEncoding3 = this.EncodingSingleton.GetEncoding_3(pCodePage, EncoderFallback, DecoderFallback)
End Function

'@Description("Returns the encoding associated with the specified code page name. Parameters specify an error handler for characters that cannot be encoded and byte sequences that cannot be decoded.")
Public Function GetEncoding4(ByVal name As String, ByVal EncoderFallback As Object, ByVal DecoderFallback As Object) As DotNetLib.Encoding
Attribute GetEncoding4.VB_Description = "Returns the encoding associated with the specified code page name. Parameters specify an error handler for characters that cannot be encoded and byte sequences that cannot be decoded."
   Set GetEncoding4 = this.EncodingSingleton.GetEncoding_4(name, EncoderFallback, DecoderFallback)
End Function

'Todo : GetEncodings
''@Description("Returns an array that contains all encodings.")
'Public Function GetEncodings() As Object()
'   GetEncodings = this.DotNetLibEncodingSingleton.GetEncodings()
'End Function

'@Description("Registers an encoding provider.")
Public Sub RegisterProvider(ByVal provider As DotNetLib.EncodingProvider)
Attribute RegisterProvider.VB_Description = "Registers an encoding provider."
    Call this.EncodingSingleton.RegisterProvider(provider)
End Sub
