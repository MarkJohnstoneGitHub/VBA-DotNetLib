VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Strings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents text as a sequence of UTF-16 code units."
'@Folder("VBADotNetLib.System")
'@PredeclaredId
'@ModuleDescription("Represents text as a sequence of UTF-16 code units.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 23, 2023
'@LastModified September 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1

'@Remarks
' Singleton class for String static members and factory methods

Option Explicit

Private Type TString
    StringStatic As DotNetLib.StringSingleton
End Type

Private this As TString

Private Sub Class_Initialize()
    Set this.StringStatic = New DotNetLib.StringSingleton
End Sub

Private Sub Class_Terminate()
    Set this.StringStatic = Nothing
End Sub

'@Static
'@Description("Replaces the format item in a specified string with the string representation of a corresponding object in a specified array.")
Public Function Format(ByVal pFormat As String, ParamArray args() As Variant) As String
Attribute Format.VB_Description = "Replaces the format item in a specified string with the string representation of a corresponding object in a specified array."
    Dim params() As Variant
    params = args
    Format = this.StringStatic.Format(pFormat, params)
End Function

'@Static
'@Description("Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information.")
Public Function Format2(ByVal provider As IFormatProvider, ByRef pFormat As String, ParamArray args() As Variant) As String
Attribute Format2.VB_Description = "Replaces the format items in a string with the string representations of corresponding objects in a specified array. A parameter supplies culture-specific formatting information."
    Dim params() As Variant
    params = args
    Format2 = this.StringStatic.Format2(provider, pFormat, params)
End Function
