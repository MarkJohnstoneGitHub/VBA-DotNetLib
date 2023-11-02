VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Booleans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a Boolean (true or false) value."
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@Folder("VBADotNetLib.System")
'@ModuleDescription("Represents a Boolean (true or false) value.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 2, 2023
'@LastModified November 2, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.boolean?view=netframework-4.8.1

Option Explicit

Const BooleanType As String = "System.Boolean"

Private Type TBoolean
    BooleanType As DotNetLib.Type
End Type

Private this As TBoolean

Private Sub Class_Initialize()
    Set this.BooleanType = Types.GetType(BooleanType)
End Sub

'--------------------------------------------------------------------
'Extension
'Added to get the type for arrays etc.
'--------------------------------------------------------------------
'@Static
'@DefaultMember
'@Description("Gets the Boolean Type.")
Public Function GetType() As DotNetLib.Type
Attribute GetType.VB_Description = "Gets the Boolean Type."
Attribute GetType.VB_UserMemId = 0
    Set GetType = this.BooleanType
End Function
