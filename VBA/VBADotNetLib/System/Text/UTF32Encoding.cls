VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UTF32Encoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a UTF-32 encoding of Unicode characters."
'@Folder "VBADotNetLib.System.Text"
'@ModuleDescription("Represents a UTF-32 encoding of Unicode characters.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 23, 2023
'@Version v1.0 November 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.utf32encoding?view=netframework-4.8.1

Option Explicit

Private Type TUTF32EncodingSingleton
   UTF32EncodingSingleton As DotNetLib.UTF32EncodingSingleton
End Type

Private this As TUTF32EncodingSingleton

Private Sub Class_Initialize()
   Set this.UTF32EncodingSingleton = New DotNetLib.UTF32EncodingSingleton
End Sub

Private Sub Class_Terminate()
   Set this.UTF32EncodingSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the UTF32Encoding class. Parameters specify whether to use the big endian byte order, whether to provide a Unicode byte order mark, and whether to throw an exception when an invalid encoding is detected.")
Public Function Create(Optional ByVal bigEndian As Boolean = False, Optional ByVal byteOrderMark As Boolean = True, Optional ByVal throwOnInvalidCharacters As Boolean = False) As DotNetLib.UTF32Encoding
Attribute Create.VB_Description = "Initializes a new instance of the UTF32Encoding class. Parameters specify whether to use the big endian byte order, whether to provide a Unicode byte order mark, and whether to throw an exception when an invalid encoding is detected."
   Set Create = this.UTF32EncodingSingleton.Create(bigEndian, byteOrderMark, throwOnInvalidCharacters)
End Function
