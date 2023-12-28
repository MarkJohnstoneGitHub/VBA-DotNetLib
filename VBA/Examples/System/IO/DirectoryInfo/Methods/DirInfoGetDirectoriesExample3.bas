Attribute VB_Name = "DirInfoGetDirectoriesExample3"
'@Folder("Examples.System.IO.DirectoryInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 26, 2023
'@LastModified December 26, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.getdirectories?view=netframework-4.8.1#system-io-directoryinfo-getdirectories(system-string-system-io-searchoption)

Option Explicit

''
' The following example lists all of the directories and files that begin with
' the letter "c" in "c:\".
''
Public Sub DirectoryInfoGetDirectoriesExample3()
    ' Specify the directory you want to manipulate.
    Dim pvtPath As String
    pvtPath = "c:\"
    Dim pvtSearchPattern As String
    pvtSearchPattern = "c*"
    
    Dim di As DotNetLib.DirectoryInfo
    Set di = DirectoryInfo.Create(pvtPath)
    
    Dim pvtDirectories() As DotNetLib.DirectoryInfo
    pvtDirectories = di.GetDirectories(pvtSearchPattern, SearchOption.SearchOption_TopDirectoryOnly)
    
    Dim pvtFiles() As DotNetLib.FileInfo
    pvtFiles = di.GetFiles(pvtSearchPattern, SearchOption.SearchOption_TopDirectoryOnly)
    
    Debug.Print VBAString.Format("Directories that begin with the letter ""c"" in {0}", pvtPath)
    Dim varDir As Variant
    For Each varDir In pvtDirectories
        Dim pvtDir As DotNetLib.DirectoryInfo
        Set pvtDir = varDir
        Debug.Print VBAString.Format("{0,-25} {1,25}", pvtDir.FullName, pvtDir.lastWriteTime)
    Next

    Debug.Print
    Debug.Print VBAString.Format("Files that begin with the letter ""c"" in {0}", pvtPath)
    Dim varFile As Variant
    For Each varFile In pvtFiles
        Dim pvtFile As DotNetLib.FileInfo
        Set pvtFile = varFile
        Debug.Print VBAString.Format("{0,-25} {1,25}", pvtFile.Name, pvtFile.lastWriteTime)
    Next
End Sub
