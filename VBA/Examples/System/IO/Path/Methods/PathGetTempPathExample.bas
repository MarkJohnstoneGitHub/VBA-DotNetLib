Attribute VB_Name = "PathGetTempPathExample"
'@Folder "Examples.System.IO.Path.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 20, 2023
'@LastModified November 20, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.gettemppath?view=netframework-4.8.1&tabs=windows#examples

Option Explicit

''
' The following code shows how to call the GetTempPath method.
''
Public Sub PathGetTempPathExample()
    Dim result As String
    result = Path.GetTempPath()
    Debug.Print result
End Sub

' This example produces output similar to the following.
' C:\Users\UserName\AppData\Local\Temp\
