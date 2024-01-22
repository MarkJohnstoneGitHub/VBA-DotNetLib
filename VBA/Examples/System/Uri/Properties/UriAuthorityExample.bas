Attribute VB_Name = "UriAuthorityExample"
'@Folder("Examples.System.Uri.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 23, 2023
'@LastModified January 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.authority?view=netframework-4.8.1#examples

Option Explicit

Public Sub UriAuthorityExample()
    Dim baseUri As DotNetLib.Uri
    Set baseUri = Uri.Create("http://www.contoso.com:8080/")
    Dim myUri As DotNetLib.Uri
    Set myUri = Uri.Create2(baseUri, "shownew.htm?date=today")
    Debug.Print myUri.Authority
End Sub

' Output:
' www.contoso.com:8080
