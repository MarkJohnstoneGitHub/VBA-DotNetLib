Attribute VB_Name = "DirectoryEnumerateFilesEg1"
'@Folder("Examples.System.IO.Directory.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 12, 2023
'@LastModified November 12, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.enumeratefiles?view=netframework-4.8.1#system-io-directory-enumeratefiles(system-string)

Option Explicit

''
' The following example shows how to retrieve all the files in a directory and
' move them to a new directory. After the files are moved, they no longer exist
' in the original directory.
''
Public Sub DirectoryEnumerateFilesExample1()
    Dim sourceDirectory As DotNetLib.String
    Set sourceDirectory = Strings.Create("C:\current")
    Dim archiveDirectory As String
    archiveDirectory = "C:\archive"
    On Error GoTo ErrorHandler
    
    Dim txtFiles As mscorlib.IEnumerable
    Set txtFiles = Directory.EnumerateFiles(sourceDirectory)
    Dim varCurrentFile As Variant
    For Each varCurrentFile In txtFiles
        Dim currentFile As DotNetLib.String
        Set currentFile = Strings.Create(varCurrentFile)
        Dim fileName As String
        fileName = currentFile.Substring(sourceDirectory.length + 1).ToString
        Call Directory.Move(currentFile.ToString, Path.Combine2(archiveDirectory, fileName))
    Next
    Exit Sub
ErrorHandler:
    Debug.Print Err.Description
End Sub
