Attribute VB_Name = "DirectoryExample"
'@Folder "Examples.System.IO.Directory"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 7, 2023
'@LastModified November 7, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory?view=netframework-4.8.1#examples

Option Explicit

''
' The following example shows how to retrieve all the text files from a
' directory and move them to a new directory. After the files are moved,
' they no longer exist in the original directory.
''
Public Sub DirectoryExample()
    Dim SourceDirectory As String
    SourceDirectory = "C:\current"
    Dim archiveDirectory As String
    archiveDirectory = "C:\archive"
    On Error Resume Next
    Dim txtFiles As mscorlib.IEnumerable
    Set txtFiles = Directory.EnumerateFiles(SourceDirectory, "*.bas")
    Dim varCurrentFile As Variant
    For Each varCurrentFile In txtFiles
        Dim fileName As String
        fileName = Mid$(varCurrentFile, Len(SourceDirectory) + 2)
        Call Directory.Move(varCurrentFile, Path.Combine2(archiveDirectory, fileName))
    Next
    If Err.Number Then
        Debug.Print "Exception: " + Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
End Sub

''
' The following example demonstrates how to use the EnumerateFiles method to
' retrieve a collection of text files from a directory, and then use that
' collection in a query to find all the lines that contain "Example".
''
Public Sub DirectoryExample2()
    Dim findText As String
    findText = "Example"
    Dim archiveDirectory As String
    archiveDirectory = "C:\archive"
    Dim txtFiles As mscorlib.IEnumerable
    Set txtFiles = Directory.EnumerateFiles(archiveDirectory, "*.bas")
    Dim foundLinesCounter As Long
    Dim varRetrievedFile As Variant
    For Each varRetrievedFile In txtFiles
        Dim fileLines As mscorlib.IEnumerable
        Set fileLines = File.ReadLines(varRetrievedFile)
        Dim i As Long
        i = 1
        Dim varfileLine As Variant
        For Each varfileLine In fileLines
            If InStr(varfileLine, findText) Then
                Debug.Print VBString.Format("{0} contains ""{1}"" at line number {2}", varRetrievedFile, findText, i)
                foundLinesCounter = foundLinesCounter + 1
            End If
            i = i + 1
        Next
    Next
    Debug.Print VBString.Format("{0} lines found.", foundLinesCounter);
End Sub

''
' The following example demonstrates how to move a directory and all its files to
' a new directory. The original directory no longer exists after it has been moved.
''
Public Sub DirectoryExample3()
    Dim SourceDirectory As String
    SourceDirectory = "C:\source"
    Dim destinationDirectory As String
    destinationDirectory = "C:\destination"
    On Error Resume Next
    Call Directory.Move(SourceDirectory, destinationDirectory)
    If Err.Number Then
        Debug.Print "Exception: " + Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
End Sub

