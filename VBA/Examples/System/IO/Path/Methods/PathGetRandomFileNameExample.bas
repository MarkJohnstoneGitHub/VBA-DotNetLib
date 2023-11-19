Attribute VB_Name = "PathGetRandomFileNameExample"
'@Folder("Examples.System.IO.Path.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 20, 2023
'@LastModified November 20, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.getrandomfilename?view=netframework-4.8.1#examples

Option Explicit

''
' The following example show output from the GetRandomFileName method.
''
Public Sub PathGetRandomFileNameExample()
    Dim result As String
    result = Path.GetRandomFileName()
    Debug.Print "Random file name is " & result
End Sub

'/*
'
' This code produces output similar to the following:
'
' Random file name is w143kxnu.idj
' Press any key to continue . . .
'
' */
