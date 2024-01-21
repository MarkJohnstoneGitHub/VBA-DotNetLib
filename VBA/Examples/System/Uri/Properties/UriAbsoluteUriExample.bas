Attribute VB_Name = "UriAbsoluteUriExample"
'@Folder("Examples.System.Uri.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 21, 2023
'@LastModified January 21, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.absoluteuri?view=netframework-4.8.1#examples

Option Explicit

''
' The following example writes the complete contents of the Uri instance to the
' console. In the example shown, http://www.contoso.com/catalog/shownew.htm?date=today
' is written to the console.
''
Public Sub UriAbsoluteUriExample()
    Dim baseUri As DotNetLib.Uri
    Set baseUri = Uri.Create("http://www.contoso.com/")
    Dim myUri As DotNetLib.Uri
    Set myUri = Uri.Create2(baseUri, "catalog/shownew.htm?date=today")
    Debug.Print myUri.AbsoluteUri
End Sub
