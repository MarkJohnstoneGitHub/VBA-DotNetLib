Attribute VB_Name = "UriAbsolutePathExample"
'@Folder("Examples.System.Uri.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 21, 2023
'@LastModified January 21, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.absolutepath?view=netframework-4.8.1

Option Explicit

''
' The following example writes the path /catalog/shownew.htm to the console.
''
Public Sub UriAbsolutePathExample()
    Dim baseUri As DotNetLib.Uri
    Set baseUri = Uri.Create("http://www.contoso.com/")
    Dim myUri As DotNetLib.Uri
    Set myUri = Uri.Create2(baseUri, "catalog/shownew.htm?date=today")
    Debug.Print myUri.AbsolutePath
End Sub
