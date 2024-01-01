VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypeCodeHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "TypeCode helper class to convert a TypeCode to string value."
'@ModuleDescription("TypeCode helper class to convert a TypeCode to string value.")
'@Folder "VBADotNetLib.System"
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 30, 2023
'@LastModified December 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.typecode?view=netframework-4.8.1

Option Explicit

Private Type TTypeCode
    TypeCodeSingleton As DotNetLib.TypeCodeHelperSingleton
End Type

Private this As TTypeCode

Private Sub Class_Initialize()
    Set this.TypeCodeSingleton = New DotNetLib.TypeCodeHelperSingleton
End Sub

Private Sub Class_Terminate()
    Set this.TypeCodeSingleton = Nothing
End Sub

'@DefaultMember
'@Description("Converts the value of a TypeCode to string value.")
Public Property Get ToString(ByVal value As mscorlib.TypeCode) As String
Attribute ToString.VB_Description = "Converts the value of a TypeCode to string value."
Attribute ToString.VB_UserMemId = 0
    ToString = this.TypeCodeSingleton.ToString(value)
End Property

'@Description("Converts the value of a TypeCode to its equivalent string representation using the specified format.")
Public Function ToString2(ByVal value As mscorlib.TypeCode, ByVal pFormat As String) As String
Attribute ToString2.VB_Description = "Converts the value of a TypeCode to its equivalent string representation using the specified format."
    ToString2 = this.TypeCodeSingleton.ToString2(value, pFormat)
End Function
