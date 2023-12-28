Attribute VB_Name = "DirInfoGetDirectoriesExample"
'@Folder("Examples.System.IO.DirectoryInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 26, 2023
'@LastModified December 26, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.getdirectories?view=netframework-4.8.1#system-io-directoryinfo-getdirectories

Option Explicit

''
' The following example retrieves all the directories in the root directory and
' displays the directory names.
''
Public Sub DirectoryInfoGetDirectoriesExample()
    Dim di As DotNetLib.DirectoryInfo
    Set di = DirectoryInfo.Create("c:\\")
    
    ' Get a reference to each directory in that directory.
    Dim diArr() As DotNetLib.DirectoryInfo
    diArr = di.GetDirectories()
    
    Dim vardri As Variant
    ' Display the names of the directories.
    For Each vardri In diArr
        Dim dri As DotNetLib.DirectoryInfo
        Set dri = vardri
        Debug.Print dri.Name
    Next
End Sub
