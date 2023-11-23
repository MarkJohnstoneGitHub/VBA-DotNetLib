VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnicodeEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a UTF-16 encoding of Unicode characters."
'@Folder "VBADotNetLib.System.Text"
'@ModuleDescription("Represents a UTF-16 encoding of Unicode characters.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 23, 2023
'@Version v1.0 November 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.unicodeencoding?view=netframework-4.8.1

Option Explicit

Private Type TUnicodeEncodingSingleton
   UnicodeEncodingSingleton As DotNetLib.UnicodeEncodingSingleton
End Type

Private this As TUnicodeEncodingSingleton

Private Sub Class_Initialize()
   Set this.UnicodeEncodingSingleton = New DotNetLib.UnicodeEncodingSingleton
End Sub

Private Sub Class_Terminate()
   Set this.UnicodeEncodingSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the UnicodeEncoding class. Parameters specify whether to use the big endian byte order, whether to provide a Unicode byte order mark, and whether to throw an exception when an invalid encoding is detected.")
Public Function Create(Optional ByVal bigEndian As Boolean = False, Optional ByVal byteOrderMark As Boolean = True, Optional ByVal throwOnInvalidBytes As Boolean = False) As DotNetLib.UnicodeEncoding
Attribute Create.VB_Description = "Initializes a new instance of the UnicodeEncoding class. Parameters specify whether to use the big endian byte order, whether to provide a Unicode byte order mark, and whether to throw an exception when an invalid encoding is detected."
   Set Create = this.UnicodeEncodingSingleton.Create(bigEndian, byteOrderMark, throwOnInvalidBytes)
End Function

'@Description("Represents the Unicode character size in bytes. This field is a constant.")
Public Property Get CharSize() As Long
Attribute CharSize.VB_Description = "Represents the Unicode character size in bytes. This field is a constant."
   CharSize = this.UnicodeEncodingSingleton.CharSize()
End Property
