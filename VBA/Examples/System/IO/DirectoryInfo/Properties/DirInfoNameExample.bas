Attribute VB_Name = "DirInfoNameExample"
'@Folder "Examples.System.IO.DirectoryInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 23, 2023
'@LastModified December 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.name?view=netframework-4.8.1#examples

Option Explicit

''
' The following example displays the name of the current DirectoryInfo instance only.
''
Public Sub DirectoryInfoNameExample()
    Dim pvtDir  As DotNetLib.DirectoryInfo
    Set pvtDir = DirectoryInfo.Create(".")
    Dim dirName As String
    dirName = pvtDir.Name
    Debug.Print VBString.Format("DirectoryInfo name is {0}.", dirName)
End Sub

' Output
' DirectoryInfo name is Documents.
