Attribute VB_Name = "UriCheckSchemeNameExample"
'@Folder "Examples.System.Uri.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 25, 2023
'@LastModified January 25, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.checkschemename?view=netframework-4.8.1#examples

Option Explicit

''
' The following example creates a Uri instance and checks whether the scheme
' name is valid.
''
Public Sub UriCheckSchemeNameExample()
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


