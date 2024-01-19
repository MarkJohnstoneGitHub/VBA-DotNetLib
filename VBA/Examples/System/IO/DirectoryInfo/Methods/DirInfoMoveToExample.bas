Attribute VB_Name = "DirInfoMoveToExample"
'@Folder "Examples.System.IO.DirectoryInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 26, 2023
'@LastModified December 26, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.moveto?view=netframework-4.8.1#examples

Option Explicit

''
' The following example demonstrates moving a directory.
''
Public Sub DirectoryInfoMoveToExample()
    ' Make a reference to a directory.
    Dim di As DotNetLib.DirectoryInfo
    Set di = DirectoryInfo.Create("TempDir")
    
    ' Create the directory only if it does not already exist.
    If (di.Exists = False) Then
        Call di.Create
    End If
    
    ' Create a subdirectory in the directory just created.
    Dim dis As DotNetLib.DirectoryInfo
    Set dis = di.CreateSubdirectory("SubDir")
    
    ' Move the main directory. Note that the contents move with the directory.
    If (Directory.Exists("NewTempDir") = False) Then
        Call di.MoveTo("NewTempDir")
    End If
    
    On Error Resume Next
    ' Attempt to delete the subdirectory. Note that because it has been
    ' moved, an exception is thrown.
    Call dis.Delete(True)
    If Err.Number Then
        ' Handle this exception in some way, such as with the following code:
        Debug.Print "That directory does not exist."
    End If
    On Error GoTo 0 'Stop code and display error
    ' Point the DirectoryInfo reference to the new directory.
    'Set di = DirectoryInfo.Create("NewTempDir")

    ' Delete the directory.
    'Call di.Delete(True)
End Sub
