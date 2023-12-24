Attribute VB_Name = "DIEnumerateDirectoriesExample"
'@Folder("Examples.System.IO.DirectoryInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 24, 2023
'@LastModified December 24, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.enumeratedirectories?view=netframework-4.8.1#system-io-directoryinfo-enumeratedirectories

Option Explicit

''
' The following example enumerates the subdirectories under the MyDocuments directory.
''
Public Sub DirectoryInfoEnumerateDirectoriesExample()
    ' Set a variable to the Documents path.
    Dim docPath As String
    docPath = Environment.GetFolderPath(SpecialFolder.SpecialFolder_MyDocuments)
    
    Dim dirPrograms As DotNetLib.DirectoryInfo
    Set dirPrograms = DirectoryInfo.Create(docPath)
    Dim varDirInfo As Variant
    For Each varDirInfo In dirPrograms.EnumerateDirectories()
        Dim di As DotNetLib.DirectoryInfo
        Set di = varDirInfo
        Debug.Print di.name
    Next
End Sub
