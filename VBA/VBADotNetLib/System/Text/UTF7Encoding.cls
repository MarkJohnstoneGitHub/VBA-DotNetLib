VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UTF7Encoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a UTF-7 encoding of Unicode characters."
'@Folder "VBADotNetLib.System.Text"
'@ModuleDescription("Represents a UTF-7 encoding of Unicode characters.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 23, 2023
'@Version v1.0 November 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.utf7encoding?view=netframework-4.8.1

Option Explicit

Private Type TUTF7EncodingSingleton
   UTF7EncodingSingleton As DotNetLib.UTF7EncodingSingleton
End Type

Private this As TUTF7EncodingSingleton

Private Sub Class_Initialize()
   Set this.UTF7EncodingSingleton = New DotNetLib.UTF7EncodingSingleton
End Sub

Private Sub Class_Terminate()
   Set this.UTF7EncodingSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the UTF7Encoding class. A parameter specifies whether to allow optional characters.")
Public Function Create(Optional ByVal allowOptionals As Boolean = False) As DotNetLib.UTF7Encoding
Attribute Create.VB_Description = "Initializes a new instance of the UTF7Encoding class. A parameter specifies whether to allow optional characters."
   Set Create = this.UTF7EncodingSingleton.Create(allowOptionals)
End Function
