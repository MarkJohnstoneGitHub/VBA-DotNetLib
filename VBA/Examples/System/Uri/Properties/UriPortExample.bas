Attribute VB_Name = "UriPortExample"
'@Folder("Examples.System.Uri.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 25, 2023
'@LastModified January 25, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.port?view=netframework-4.8.1#examples

Option Explicit

''
' The following example writes the URI port number to the console. In this case,
' the value is the default port number for HTTP, port 80.
''
Public Sub UriPortExample()
    Dim baseUri As DotNetLib.Uri
    Set baseUri = Uri.Create("http://www.contoso.com/")
    Dim myUri As DotNetLib.Uri
    Set myUri = Uri.Create2(baseUri, "catalog/shownew.htm?date=today")

    Debug.Print myUri.Port
End Sub

' Ouput
'   80
