Attribute VB_Name = "DirectoryMoveExample"
'@Folder "Examples.System.IO.Directory.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 19, 2023
'@LastModified November 19, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.move?view=netframework-4.8.1#examples

Option Explicit

''
' The following example demonstrates how to move a directory and all its files to
' a new directory. The original directory no longer exists after it has been moved.
''
Public Sub DirectoryMove()
    Dim SourceDirectory As String
    SourceDirectory = "C:\source"
    Dim destinationDirectory As String
    destinationDirectory = "C:\destination"
    
    On Error Resume Next
    Call Directory.Move(SourceDirectory, destinationDirectory)
    If Err.Number <> 0 Then
        Debug.Print Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
End Sub
