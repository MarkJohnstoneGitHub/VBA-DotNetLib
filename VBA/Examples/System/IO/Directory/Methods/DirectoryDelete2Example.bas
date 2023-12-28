Attribute VB_Name = "DirectoryDelete2Example"
'@Folder("Examples.System.IO.Directory.Methods")
'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 16, 2023
'@LastModified November 16, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.delete?view=netframework-4.8.1#system-io-directory-delete(system-string-system-boolean)

Option Explicit

''
' The following example shows how to create a new directory, subdirectory, and
' file in the subdirectory, and then recursively delete all the new items.
''
Public Sub DirectoryDelete2Example()
    Dim topPath As String
    topPath = "C:\NewDirectory"
    Dim subPath As String
    subPath = "C:\NewDirectory\NewSubDirectory"
    On Error GoTo ErrorHandler
    Call Directory.CreateDirectory(subPath)
    Dim writer As DotNetLib.StreamWriter
    Set writer = File.createText(subPath + "\example.txt")
    Call writer.WriteLine2("content added")
    Call writer.Dispose
    Call Directory.Delete2(topPath, True)
    Dim directoryExists As Boolean
    directoryExists = Directory.Exists(topPath)
    Debug.Print "top-level directory exists: " & directoryExists
Exit Sub
ErrorHandler:
    Debug.Print VBAString.Format("The process failed: {0}", Err.Description)
End Sub
