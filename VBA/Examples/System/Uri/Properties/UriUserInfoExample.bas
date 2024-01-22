Attribute VB_Name = "UriUserInfoExample"
'@Folder("Examples.System.Uri.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 22, 2023
'@LastModified January 22, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.userinfo?view=netframework-4.8.1#examples

Option Explicit

''
' The following example creates a Uri instance and writes the user information
' to the console.
''
Public Sub UriUserInfoExample()
    Dim uriAddress As DotNetLib.Uri
    Set uriAddress = Uri.Create("http://user:password@www.contoso.com/index.htm ")
    Debug.Print uriAddress.UserInfo
    Debug.Print VBString.Format("Fully Escaped {0}", IIf(uriAddress.UserEscaped, "yes", "no"))
End Sub

'Output
'   user:password
'   Fully Escaped no

