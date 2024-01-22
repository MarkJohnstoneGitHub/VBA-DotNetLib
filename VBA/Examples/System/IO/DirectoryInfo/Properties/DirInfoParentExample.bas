Attribute VB_Name = "DirInfoParentExample"
'@Folder "Examples.System.IO.DirectoryInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 23, 2023
'@LastModified December 24, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.parent?view=netframework-4.8.1#examples

Option Explicit

''
' The following example refers to the parent directory of a specified directory.
''
Public Sub DirectoryInfoParentExample()
    ' Make a reference to a directory.
    Dim di As DotNetLib.DirectoryInfo
    Set di = DirectoryInfo.Create("TempDir")
    
    '/ Create the directory only if it does not already exist.
    If (di.Exists = False) Then
        Call di.Create
    End If
    
    ' Create a subdirectory in the directory just created.
    Dim dis As DotNetLib.DirectoryInfo
    Set dis = di.CreateSubdirectory("SubDir")
    
    ' Get a reference to the parent directory of the subdirectory you just made.
    Dim parentDir As DotNetLib.DirectoryInfo
    Set parentDir = dis.Parent
    Debug.Print VBString.Format("The parent directory of '{0}' is '{1}'", dis.name, parentDir.name)
    
    ' Delete the parent directory.
    Call di.Delete(True)
End Sub

' Output
' The parent directory of 'SubDir' is 'TempDir'
