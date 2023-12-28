VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringComparisonHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "StringComparison helper class to convert a StringComparison to string value."
'Rubberduck annotations
'@ModuleDescription("StringComparison helper class to convert a StringComparison to string value.")
'@Folder "VBADotNetLib.System"
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 29, 2023
'@LastModified December 29, 2023

'@ReferenceAddin DotNetLib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.stringcomparison?view=netframework-4.8.1

Option Explicit

'@DefaultMember
'@Description("Converts the value of a DateTimeKind to string value.")
Public Property Get ToString(ByVal value As mscorlib.StringComparison) As String
Attribute ToString.VB_Description = "Converts the value of a DateTimeKind to string value."
Attribute ToString.VB_UserMemId = 0
    Select Case value
        Case StringComparison_CurrentCulture: ToString = "CurrentCulture"
        Case StringComparison_CurrentCultureIgnoreCase: ToString = "CurrentCultureIgnoreCase"
        Case StringComparison_InvariantCulture: ToString = "InvariantCulture"
        Case StringComparison_InvariantCultureIgnoreCase: ToString = "InvariantCultureIgnoreCase"
        Case StringComparison_Ordinal: ToString = "Ordinal"
        Case StringComparison_OrdinalIgnoreCase: ToString = "OrdinalIgnoreCase"
    End Select
End Property
