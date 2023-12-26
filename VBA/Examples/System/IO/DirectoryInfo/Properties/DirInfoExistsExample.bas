Attribute VB_Name = "DirInfoExistsExample"
'@Folder("Examples.System.IO.DirectoryInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 23, 2023
'@LastModified December 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.exists?view=netframework-4.8.1#examples

Option Explicit

''
' The following example demonstrates a use of the DirectoryInfo.Exists property
' in the context of copying a source directory to a target directory.
''
Public Sub DirectoryInfoExists()
    Call CopyDirectory("D:\\Tools", "D:\\NewTools")
End Sub

Private Sub CopyDirectory(ByVal SourceDirectory As String, ByVal TargetDirectory As String)
    Dim source As DotNetLib.DirectoryInfo
    Set source = DirectoryInfo.Create(SourceDirectory)
    Dim target As DotNetLib.DirectoryInfo
    Set target = DirectoryInfo.Create(TargetDirectory)
    
    ' Determine whether the source directory exists.
    If (Not source.Exists) Then
        Exit Sub 'Return
    End If
    
    If (Not target.Exists) Then
        Call target.Create
    End If
    ' Copy files.
    Dim sourceFiles() As DotNetLib.FileInfo
    sourceFiles = source.GetFiles()
    Dim i As Long
    For i = 0 To UBound(sourceFiles)
        Call file.Copy2(sourceFiles(i).FullName, target.FullName + "\\" + sourceFiles(i).name, True)
    Next i

    ' Copy directories.
    Dim sourceDirectories() As DotNetLib.DirectoryInfo
    sourceDirectories = source.GetDirectories()
    Dim j As Long
    For j = 0 To UBound(sourceDirectories)
        Call CopyDirectory(sourceDirectories(j).FullName, target.FullName + "\\" + sourceDirectories(j).name)
    Next j
End Sub
