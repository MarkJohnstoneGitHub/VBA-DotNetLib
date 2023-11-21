Attribute VB_Name = "FileAppendAllTextExample"
'@Folder("Examples.System.IO.File.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 21, 2023
'@LastModified November 21, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.file.appendalltext?view=netframework-4.8.1#system-io-file-appendalltext(system-string-system-string)

Option Explicit

''
' The following code example demonstrates the use of the AppendAllText method
' to add extra text to the end of a file. In this example, a file is created if
' it doesn't already exist, and text is added to it. However, the directory
' named temp on drive C must exist for the example to complete successfully.
''
Public Sub FileAppendAllTextExample()
    Dim pvtPath As String
    pvtPath = "c:\temp\MyTest.txt"
    
    ' This text is added only once to the file.
    If (Not File.Exists(pvtPath)) Then
        ' Create a file to write to.
        Dim createText As String
        createText = "Hello and Welcome" + Environment.NewLine
        Call File.WriteAllText(pvtPath, createText)
    End If
    
    ' This text is always added, making the file longer over time
    ' if it is not deleted.
    Dim appendText As String
    appendText = "This is extra text" + Environment.NewLine
    Call File.AppendAllText(pvtPath, appendText)
End Sub
