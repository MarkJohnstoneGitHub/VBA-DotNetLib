VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Int16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a 16-bit signed integer."
'@Folder("VBADotNetLib.System")
'@ModuleDescription("Represents a 16-bit signed integer.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 12, 2023
'@LastModified October 15, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.int16?view=netframework-4.8.1

'@Remarks
' Singleton class for Int16 containing constructors and static members.

Option Explicit

Const Int16Type As String = "System.Int16"

Private Type TInt16
    Int16Type       As DotNetLib.Type
End Type

Private this As TInt16

Private Sub Class_Initialize()
    Set this.Int16Type = Types.GetType(Int16Type)
End Sub

Private Sub Class_Terminate()
    Set this.Int16Type = Nothing
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
    Set GetType = this.Int16Type
End Function
