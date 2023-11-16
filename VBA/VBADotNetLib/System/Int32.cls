VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Int32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a 32-bit signed integer."
'@Folder "VBADotNetLib.System"
'@ModuleDescription("Represents a 32-bit signed integer.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 12, 2023
'@LastModified November 2, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.int32?view=netframework-4.8.1

'@Remarks
' Singleton class for Int32 containing constructors and static members.

Option Explicit

Const Int32Type As String = "System.Int32"

Private Type TInt32
    Int32Type       As DotNetLib.Type
End Type

Private this As TInt32

Private Sub Class_Initialize()
    Set this.Int32Type = Types.GetType(Int32Type)
End Sub

Private Sub Class_Terminate()
    Set this.Int32Type = Nothing
End Sub


'--------------------------------------------------------------------
'Extension
'Added to get the type for arrays etc.
'--------------------------------------------------------------------
'@Static
'@DefaultMember
'@Description("Gets the Int32 Type.")
Public Function GetType() As DotNetLib.Type
Attribute GetType.VB_Description = "Gets the Int32 Type."
Attribute GetType.VB_UserMemId = 0
    Set GetType = this.Int32Type
End Function
