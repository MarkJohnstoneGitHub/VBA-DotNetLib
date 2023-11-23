VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UTF8Encoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a UTF-8 encoding of Unicode characters."
'@Folder "VBADotNetLib.System.Text"
'@ModuleDescription("Represents a UTF-8 encoding of Unicode characters.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 23, 2023
'@Version v1.0 November 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.utf8encoding?view=netframework-4.8.1

Option Explicit

Private Type TUTF8EncodingSingleton
   UTF8EncodingSingleton As DotNetLib.UTF8EncodingSingleton
End Type

Private this As TUTF8EncodingSingleton

Private Sub Class_Initialize()
   Set this.UTF8EncodingSingleton = New DotNetLib.UTF8EncodingSingleton
End Sub

Private Sub Class_Terminate()
   Set this.UTF8EncodingSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the UTF8Encoding class. Parameters specify whether to provide a Unicode byte order mark and whether to throw an exception when an invalid encoding is detected.")
Public Function Create(Optional ByVal encoderShouldEmitUTF8Identifier As Boolean = False, Optional ByVal throwOnInvalidBytes As Boolean = False) As DotNetLib.UTF8Encoding
Attribute Create.VB_Description = "Initializes a new instance of the UTF8Encoding class. Parameters specify whether to provide a Unicode byte order mark and whether to throw an exception when an invalid encoding is detected."
   Set Create = this.UTF8EncodingSingleton.Create(encoderShouldEmitUTF8Identifier, throwOnInvalidBytes)
End Function
