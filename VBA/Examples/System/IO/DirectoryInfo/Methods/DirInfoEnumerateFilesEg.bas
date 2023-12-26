Attribute VB_Name = "DirInfoEnumerateFilesEg"
'@Folder("Examples.System.IO.DirectoryInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 26, 2023
'@LastModified December 26, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.enumeratefiles?view=netframework-4.8.1#system-io-directoryinfo-enumeratefiles

Option Explicit

''
' The following example enumerates the files under a specified directory.
''
Public Sub DirectoryInfoEnumerateFiles()
    Dim dirInfo As DotNetLib.DirectoryInfo
    Set dirInfo = DirectoryInfo.Create("\\archives1\library\")

    Dim files As mscorlib.IEnumerable
    Set files = dirInfo.EnumerateFiles()
    
    Dim varfile As Variant
    For Each varfile In files
        Dim file As DotNetLib.FileInfo
        Set file = varfile
        Debug.Print VBAString.Format("{0}", file.name)
    Next
End Sub
