Attribute VB_Name = "DirectoryInfoDeleteExample2"
'@Folder("Examples.System.IO.DirectoryInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 23, 2023
'@LastModified December 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.delete?view=netframework-4.8.1#system-io-directoryinfo-delete(system-boolean)

Option Explicit

''
' The following example demonstrates deleting a directory. Because the directory
' is removed, first comment out the Delete line to test that the directory exists.
' Then uncomment the same line of code to test that the directory was removed
' successfully.
''
Public Sub DirectoryInfoDeleteExample2()
    ' Make a reference to a directory.
    Dim di As DotNetLib.DirectoryInfo
    Set di = DirectoryInfo.Create("TempDir")
    
    ' Create the directory only if it does not already exist.
    If (di.Exists = False) Then
        Call di.Create
    End If
    
    '/ Create a subdirectory in the directory just created.
    Dim dis As DotNetLib.DirectoryInfo
    Set dis = di.CreateSubdirectory("SubDir")

    ' Process that directory as required.
    ' ...

    ' Delete the subdirectory. The true indicates that if subdirectories
    ' or files are in this directory, they are to be deleted as well.
    Call dis.Delete(True)

    ' Delete the directory.
    Call di.Delete(True)
End Sub
