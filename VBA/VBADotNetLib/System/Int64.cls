VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Int64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a 64-bit signed integer."
'@Folder("VBADotNetLib.System")
'@ModuleDescription("Represents a 64-bit signed integer.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 12, 2023
'@LastModified October 12, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.int64?view=netframework-4.8.1

'@Remarks
' Singleton class for Int64 containing constructors and static members.

Option Explicit

Const Int64Type As String = "System.Int64"

Private Type TInt64
    Int64Type       As DotNetLib.Type
End Type

Private this As TInt64

Private Sub Class_Initialize()
    Set this.Int64Type = Types.GetType(Int64Type)
End Sub

Private Sub Class_Terminate()
    Set this.Int64Type = Nothing
End Sub

'--------------------------------------------------------------------
'Extension
'Added to get the type for arrays etc.
'--------------------------------------------------------------------
'@Static
'@DefaultMember
'@Description("Gets the String Type.")
Public Function GetType() As DotNetLib.Type
    Set GetType = this.Int64Type
End Function

