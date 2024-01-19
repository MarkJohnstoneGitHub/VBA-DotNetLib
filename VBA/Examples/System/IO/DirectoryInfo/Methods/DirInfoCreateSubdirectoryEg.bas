Attribute VB_Name = "DirInfoCreateSubdirectoryEg"
'@Folder "Examples.System.IO.DirectoryInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 23, 2023
'@LastModified December 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.createsubdirectory?view=netframework-4.8.1#system-io-directoryinfo-createsubdirectory(system-string)

Option Explicit

''
' The following example demonstrates creating a subdirectory.
' In this example, the created directories are removed once created. Therefore,
' to test this sample, comment out the delete lines in the code.
''
Public Sub DirectoryInfoCreateSubdirectory()
    ' Create a reference to a directory.
    Dim di As DotNetLib.DirectoryInfo
    Set di = DirectoryInfo.Create("TempDir")
    
    ' Create the directory only if it does not already exist.
    If (di.Exists = False) Then
        Call di.Create
    End If
    
    ' Create a subdirectory in the directory just created.
    Dim dis As DotNetLib.DirectoryInfo
    Set dis = di.CreateSubdirectory("SubDir")
    
    ' Process that directory as required.
    ' ...
    
    ' Delete the subdirectory.
    Call dis.Delete(True)

    ' Delete the directory.
    Call di.Delete(True)
End Sub
