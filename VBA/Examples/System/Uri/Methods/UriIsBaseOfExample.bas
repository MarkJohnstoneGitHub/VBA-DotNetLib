Attribute VB_Name = "UriIsBaseOfExample"
'@Folder "Examples.System.Uri.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 26, 2023
'@LastModified January 26, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.isbaseof?view=netframework-4.8.1#examples

Option Explicit

''
' This example creates a Uri instance that represents a base Uri instance.
' It then creates a second Uri instance from a string. It calls IsBaseOf to
' determine whether the base instance is the base of the second instance.
' The outcome is written to the console.
''
Public Sub UriIsBaseOfExample()
    ' Create a base Uri.
    Dim baseUri As DotNetLib.Uri
    Set baseUri = Uri.Create("http://www.contoso.com/")
    
    ' Create a new Uri from a string.
    Dim uriAddress As DotNetLib.Uri
    Set uriAddress = Uri.Create("http://www.contoso.com/index.htm?date=today")
    
    ' Determine whether BaseUri is a base of UriAddress.
    If (baseUri.IsBaseOf(uriAddress)) Then
        Debug.Print VBString.Format("{0} is the base of {1}", baseUri, uriAddress)
    End If
End Sub

' Output
'   http://www.contoso.com/ is the base of http://www.contoso.com/index.htm?date=today
