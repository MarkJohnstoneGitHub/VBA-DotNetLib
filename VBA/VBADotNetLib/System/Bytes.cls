VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Bytes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents an 8-bit unsigned integer."
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@Folder("VBADotNetLib.System")
'@ModuleDescription("Represents an 8-bit unsigned integer.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 2, 2023
'@LastModified November 2, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.boolean?view=netframework-4.8.1

Option Explicit

Const ByteType As String = "System.Byte"

Private Type TByte
    ByteType As DotNetLib.Type
End Type

Private this As TByte

Private Sub Class_Initialize()
    Set this.ByteType = Types.GetType(ByteType)
End Sub

'--------------------------------------------------------------------
'Extension
'Added to get the type for arrays etc.
'--------------------------------------------------------------------
'@Static
'@DefaultMember
'@Description("Gets the Byte Type.")
Public Function GetType() As DotNetLib.Type
Attribute GetType.VB_Description = "Gets the Byte Type."
Attribute GetType.VB_UserMemId = 0
    Set GetType = this.ByteType
End Function
