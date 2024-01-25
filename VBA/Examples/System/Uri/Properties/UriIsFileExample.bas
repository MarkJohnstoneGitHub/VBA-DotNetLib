Attribute VB_Name = "UriIsFileExample"
'@Folder("Examples.System.Uri.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 25, 2023
'@LastModified January 25, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.isfile?view=netframework-4.8.1#examples

Option Explicit

''
' The following example creates a Uri instance and determines whether it is a
' file URI.
''
Public Sub UriIsFileExample()
    Dim uriAddress2 As DotNetLib.Uri
    Set uriAddress2 = Uri.Create("file://server/filename.ext")
    Debug.Print uriAddress2.LocalPath
    Debug.Print VBString.Format("Uri {0} a UNC path", IIf(uriAddress2.IsUnc, "is", "is not"))
    Debug.Print VBString.Format("Uri {0} a local host", IIf(uriAddress2.IsLoopback, "is", "is not"))
    Debug.Print VBString.Format("Uri {0} a file", IIf(uriAddress2.IsFile, "is", "is not"))
End Sub

' The example displays the following output:
'    \\server\filename.ext
'    Uri is a UNC path
'    Uri is not a local host
'    Uri is a file
