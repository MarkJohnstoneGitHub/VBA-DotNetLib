Attribute VB_Name = "FileAppendAllText2Example"
'@Folder "Examples.System.IO.File.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 23, 2023
'@LastModified November 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.file.appendalltext?view=netframework-4.8.1#system-io-file-appendalltext(system-string-system-string-system-text-encoding)

Option Explicit

''
' The following code example demonstrates the use of the AppendAllText method
' to add extra text to the end of a file. In this example, a file is created
' if it doesn't already exist, and text is added to it. However, the directory
' named temp on drive C must exist for the example to complete successfully.
''
Public Sub FileAppendAllText2Example()
    Dim pvtPath As String
    pvtPath = "c:\temp\MyTest.txt"
    
    ' This text is added only once to the file.
    If (Not File.Exists(pvtPath)) Then
        ' Create a file to write to.
        Dim createText As String
        createText = "Hello and Welcome" + Environment.NewLine
        Call File.WriteAllText2(pvtPath, createText, Encoding.UTF8)
    End If
    
    ' This text is always added, making the file longer over time
    ' if it is not deleted.
    Dim pvtAppendText As String
    pvtAppendText = "This is extra text" + Environment.NewLine
    Call File.AppendAllText2(pvtPath, pvtAppendText, Encoding.UTF8)
End Sub

