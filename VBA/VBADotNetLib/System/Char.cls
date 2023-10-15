VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Char"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBADotNetLib.System")
'@ModuleDescription("Represents a character as a UTF-16 code unit.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 15, 2023
'@LastModified October 15, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.char?view=netframework-4.8.1

'@Remarks
' Singleton class for Char containing constructors and static members.

Option Explicit

Const CharType As String = "System.Char"

Private Type TChar
    CharType As DotNetLib.Type
End Type

Private this As TChar

Private Sub Class_Initialize()
    Set this.CharType = Types.GetType(CharType)
End Sub

Private Sub Class_Terminate()
    Set this.CharType = Nothing
End Sub

'--------------------------------------------------------------------
'Extension
'Added to get the type for arrays etc.
'--------------------------------------------------------------------
'@Static
'@DefaultMember
'@Description("Gets the String Type.")
Public Function GetType() As DotNetLib.Type
Attribute GetType.VB_Description = "Gets the String Type."
Attribute GetType.VB_UserMemId = 0
    Set GetType = this.CharType
End Function
