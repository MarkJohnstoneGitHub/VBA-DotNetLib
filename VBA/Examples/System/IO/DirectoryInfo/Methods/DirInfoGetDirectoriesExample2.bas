Attribute VB_Name = "DirInfoGetDirectoriesExample2"
'@Folder "Examples.System.IO.DirectoryInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 26, 2023
'@LastModified December 26, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.getdirectories?view=netframework-4.8.1#system-io-directoryinfo-getdirectories(system-string)

Option Explicit

''
' The following example counts the directories in a path that contain the
'specified letter.
''
Public Sub DirectoryInfoGetDirectoriesExample2()
    On Error GoTo ErrorHandler
    Dim di As DotNetLib.DirectoryInfo
    Set di = DirectoryInfo.Create("c:\\")
    
    ' Get only subdirectories that contain the letter "p."
    Dim dirs() As DotNetLib.DirectoryInfo
    dirs = di.GetDirectories("*p*")
    Debug.Print VBString.Format("The number of directories containing the letter p is {0}.", UBound(dirs) + 1)
    
    Dim vardiNext As Variant
    For Each vardiNext In dirs
        Dim diNext As DotNetLib.DirectoryInfo
        Set diNext = vardiNext
        Debug.Print VBString.Format("The number of files in {0} is {1}", diNext, UBound(diNext.GetFiles()) + 1)
    Next
    Exit Sub
    
ErrorHandler:
    Debug.Print VBString.Format("The process failed: {0}", Err.Description)
End Sub
