VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASCIIEncoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents an ASCII character encoding of Unicode characters."
'@Folder "VBADotNetLib.System.Text"
'@ModuleDescription("Represents an ASCII character encoding of Unicode characters.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 23, 2023
'@Version v1.0 November 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.asciiencoding?view=netframework-4.8.1

Option Explicit

Private Type TASCIIEncodingSingleton
   ASCIIEncodingSingleton As DotNetLib.ASCIIEncodingSingleton
End Type

Private this As TASCIIEncodingSingleton

Private Sub Class_Initialize()
   Set this.ASCIIEncodingSingleton = New DotNetLib.ASCIIEncodingSingleton
End Sub

Private Sub Class_Terminate()
   Set this.ASCIIEncodingSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the ASCIIEncoding class.")
Public Function Create() As DotNetLib.ASCIIEncoding
Attribute Create.VB_Description = "Initializes a new instance of the ASCIIEncoding class."
   Set Create = this.ASCIIEncodingSingleton.Create()
End Function
