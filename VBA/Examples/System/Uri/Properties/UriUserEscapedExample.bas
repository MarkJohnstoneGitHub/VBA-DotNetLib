Attribute VB_Name = "UriUserEscapedExample"
'@Folder("Examples.System.Uri.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 25, 2023
'@LastModified January 25, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.userescaped?view=netframework-4.8.1#examples

Option Explicit

''
' The following example creates a Uri instance and determines whether it was
' fully escaped when it was created.
''
Public Sub UriUserEscapedExample()
    Dim uriAddress As DotNetLib.Uri
    Set uriAddress = Uri.Create("http://user:password@www.contoso.com/index.htm ")
    Debug.Print uriAddress.UserInfo
    Debug.Print VBString.Format("Fully Escaped {0}", IIf(uriAddress.UserEscaped, "yes", "no"))
End Sub

' Ouput
'   user:password
'   Fully Escaped no

