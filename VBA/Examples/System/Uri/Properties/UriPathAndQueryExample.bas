Attribute VB_Name = "UriPathAndQueryExample"
'@Folder("Examples.System.Uri.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 23, 2023
'@LastModified January 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.pathandquery?view=netframework-4.8.1#examples

Option Explicit

''
' The following example writes the URI path (/catalog/shownew.htm) and query
' (?date=today) information to the console.
''
Public Sub UriPathAndQueryExample()
    Dim baseUri As DotNetLib.Uri
    Set baseUri = Uri.Create("http://www.contoso.com/")
    Dim myUri As DotNetLib.Uri
    Set myUri = Uri.Create2(baseUri, "catalog/shownew.htm?date=today")
    Debug.Print myUri.PathAndQuery
End Sub

' Output
' /catalog/shownew.htm?date=today
