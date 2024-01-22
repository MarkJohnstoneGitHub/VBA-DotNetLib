Attribute VB_Name = "UriHostNameTypeExample"
'@Folder("Examples.System.Uri.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 23, 2023
'@LastModified January 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.hostnametype?view=netframework-4.8.1#examples

Option Explicit

''
' The following example creates a Uri instance and writes the HostNameType to
' the console.
''
Public Sub UriHostNameTypeExample()
    Dim address1 As DotNetLib.Uri
    Set address1 = Uri.Create("http://www.contoso.com/index.htm#search")
    Debug.Print VBString.Format("address 1 {0} a valid scheme name", _
                IIf(Uri.CheckSchemeName(address1.Scheme), " has", " does not have"))

    If (address1.Scheme = Uri.UriSchemeHttp) Then
        Debug.Print "Uri is HTTP type"
    End If

    Debug.Print UriHostNameTypeHelper.ToString(address1.HostNameType)
End Sub

' Output
'    address 1  has a valid scheme name
'    Uri is HTTP type
'    Dns
