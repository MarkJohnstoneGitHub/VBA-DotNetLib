Attribute VB_Name = "DirectoryGetDirectoriesExample3"
'@Folder "Examples.System.IO.Directory.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 18, 2023
'@LastModified November 18, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.getdirectories?view=netframework-4.8.1#system-io-directory-getdirectories(system-string-system-string-system-io-searchoption)

Option Explicit

''
' The following example counts the number of directories that begin with the
' specified letter in a path. Only the top-level directory is searched.
''
Public Sub DirectoryGetDirectoriesExample3()
    On Error GoTo ErrorHandler

    Dim dirs() As String
    dirs = Directory.GetDirectories("c:\", "p*", SearchOption.SearchOption_TopDirectoryOnly)
    Debug.Print VBString.Format("The number of directories starting with p is {0}.", UBound(dirs) + 1)
    Dim dir As Variant
    For Each dir In dirs
        Debug.Print dir
    Next
Exit Sub
ErrorHandler:
    Debug.Print VBString.Format("The process failed: {0}", Err.Description)
End Sub
