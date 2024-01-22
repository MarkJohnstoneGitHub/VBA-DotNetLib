Attribute VB_Name = "UriHostExample"
'@Folder("Examples.System.Uri.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 23, 2023
'@LastModified January 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.host?view=netframework-4.8.1#examples

Option Explicit

''
' The following example writes the host name (www.contoso.com) of the server
' to the console.
''
Public Sub UriHostExample()
    Dim baseUri As DotNetLib.Uri
    Set baseUri = Uri.Create("http://www.contoso.com:8080/")
    Dim myUri As DotNetLib.Uri
    Set myUri = Uri.Create2(baseUri, "shownew.htm?date=today")

    Debug.Print myUri.Host
End Sub

' Output
'   www.contoso.com
