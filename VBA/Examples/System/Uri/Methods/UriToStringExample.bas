Attribute VB_Name = "UriToStringExample"
'@Folder "Examples.System.Uri.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 26, 2023
'@LastModified January 26, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.tostring?view=netframework-4.8.1#examples

Option Explicit

''
' The following example creates a new Uri instance from a string. It illustrates
' the difference between the value returned from OriginalString, which returns
' the string that was passed to the constructor, and from a call to ToString,
' which returns the canonical form of the string.
''
Public Sub UriToStringExample()
    ' Create a new Uri from a string address.
    Dim uriAddress As DotNetLib.Uri
    Set uriAddress = Uri.Create("HTTP://www.Contoso.com:80/thick%20and%20thin.htm")
    
    ' Write the new Uri to the console and note the difference in the two values.
    ' ToString() gives the canonical version.  OriginalString gives the orginal
    ' string that was passed to the constructor.
    
    ' The following outputs "http://www.contoso.com/thick and thin.htm".
    Debug.Print uriAddress.ToString()
    
    ' The following outputs "HTTP://www.Contoso.com:80/thick%20and%20thin.htm".
    Debug.Print uriAddress.OriginalString
End Sub

' Output
'    http://www.contoso.com/thick and thin.htm
'    HTTP://www.Contoso.com:80/thick%20and%20thin.htm
