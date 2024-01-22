VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UriHostNameTypeHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "UriHostNameType helper class to convert a UriHostNameType enum to string value."
'@Folder("VBADotNetLib.System")
'@PredeclaredId
'@ModuleDescription("UriHostNameType helper class to convert a UriHostNameType enum to string value.")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 23, 2023
'@LastModified January 23, 2023

'@ReferenceAddin DotNetLib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.urihostnametype?view=netframework-4.8.1

Option Explicit

'@DefaultMember
'@Description("Converts the value of a UriHostNameType to string value.")
Public Property Get ToString(ByVal value As DotNetLib.UriHostNameType) As String
Attribute ToString.VB_Description = "Converts the value of a UriHostNameType to string value."
Attribute ToString.VB_UserMemId = 0
    Select Case value
        Case UriHostNameType_Basic: ToString = "Basic"
        Case UriHostNameType_Dns: ToString = "Dns"
        Case UriHostNameType_IPv4: ToString = "IPv4"
        Case UriHostNameType_IPv6: ToString = "IPv6"
        Case UriHostNameType_Unknown: ToString = "Unknown"
    End Select
End Property
