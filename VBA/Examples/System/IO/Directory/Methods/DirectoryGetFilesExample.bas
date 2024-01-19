Attribute VB_Name = "DirectoryGetFilesExample"
'@Folder "Examples.System.IO.Directory.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 18, 2023
'@LastModified November 18, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.getfiles?view=netframework-4.8.1#system-io-directory-getfiles(system-string-system-string)

Option Explicit

''
' The following example counts the number of files that begin with the
' specified letter.
''
Public Sub DirectoryGetFilesEg()
    On Error GoTo ErrorHandler
    ' Only get files that begin with the letter "c".
    Dim dirs() As String
    dirs = Directory.GetFiles("c:\VBA\Export", "c*")
    Debug.Print VBString.Format("The number of files starting with c is {0}.", UBound(dirs) + 1)
    Dim dir As Variant
    For Each dir In dirs
        Debug.Print dir
    Next
Exit Sub
ErrorHandler:
    Debug.Print VBString.Format("The process failed: {0}", Err.Description)
End Sub
