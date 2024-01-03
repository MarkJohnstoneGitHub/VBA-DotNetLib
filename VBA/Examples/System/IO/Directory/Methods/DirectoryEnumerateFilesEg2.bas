Attribute VB_Name = "DirectoryEnumerateFilesEg2"
'@Folder("Examples.System.IO.Directory.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 14, 2023
'@LastModified December 28, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.enumeratefiles?view=netframework-4.8.1#system-io-directory-enumeratefiles(system-string-system-string)

Option Explicit

''
' The following example enumerates the files in the specified directory,
' reads each line of the file, and displays the line if it contains the
' string "DotNetLib.String".
''
Public Sub DirectoryEnumerateFilesExample2()
    On Error GoTo ErrorHandler
    Dim txtFiles As mscorlib.IEnumerable
    Set txtFiles = Directory.EnumerateFiles("C:\VBA\Export", "*.bas") 'Eg Containing export of VBADotNetLibrary

    Dim varCurrentFile As Variant
    For Each varCurrentFile In txtFiles
        Dim currentFile As DotNetLib.String
        Set currentFile = Strings.Create(varCurrentFile)
        Dim filename As String
        Dim txtFile() As String
        txtFile = File.ReadAllLines(varCurrentFile)
        Dim varLine As Variant
        Dim lineCounter As Long
        lineCounter = 0
        For Each varLine In txtFile
            lineCounter = lineCounter + 1
            Dim line As DotNetLib.String
            Set line = Strings.Create(varLine)
            If line.Contains2("DotNetLib.String") Then
                Debug.Print varCurrentFile & " contains DotNetLib.String at line " & lineCounter
            End If
        Next
    Next
    Exit Sub
ErrorHandler:
    Debug.Print Err.Description
End Sub

