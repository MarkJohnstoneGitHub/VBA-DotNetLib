Attribute VB_Name = "DirInfoGetFileSystemInfosEg"
'@Folder "Examples.System.IO.DirectoryInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 26, 2023
'@LastModified December 28, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.getfilesysteminfos?view=netframework-4.8.1

Option Explicit

Private numFiles As Long
Private numDirectories As Long

''
' The following example counts the files and directories under the specified directory.
''
Public Sub DirectoryInfoGetFileSystemInfosExample()
    numFiles = 0
    numDirectories = 0

    On Error GoTo ErrorHandler
    
    Dim pvtPath As String
    pvtPath = InputBox("Enter the path to a directory:")

    ' Create a new DirectoryInfo object.
    Dim pvtDir As DotNetLib.DirectoryInfo
    Set pvtDir = DirectoryInfo.Create(pvtPath)
    If (Not pvtDir.Exists) Then
        Err.Raise DirectoryNotFoundException, Description:="The directory does not exist."
    End If
    
    ' Call the GetFileSystemInfos method.
    Dim infos() As DotNetLib.FileSystemInfo
    infos = pvtDir.GetFileSystemInfos()

    Debug.Print "Working..."
    
    ' Pass the result to the ListDirectoriesAndFiles
    ' method defined below.
    Call ListDirectoriesAndFiles(infos)

    ' Display the results to the console.
    Debug.Print VBString.Format("Directories: {0}", numDirectories)
    Debug.Print VBString.Format("Files: {0}", numFiles)
Exit Sub
ErrorHandler:
    Debug.Print Err.Description
End Sub

Private Sub ListDirectoriesAndFiles(ByRef FSInfo() As DotNetLib.FileSystemInfo)
    '/ Check the FSInfo parameter.
    ' If (FSInfo = Null) Then
    '            throw new ArgumentNullException("FSInfo");
    ' End If

    ' Iterate through each item.
    Dim varItem As Variant
    For Each varItem In FSInfo
        Dim fileSytemInfoItem As DotNetLib.FileSystemInfo
        Set fileSytemInfoItem = varItem
        
        ' Check to see if this is a DirectoryInfo object.
        If TypeOf fileSytemInfoItem Is DotNetLib.DirectoryInfo Then
            ' Add one to the directory count.
            numDirectories = numDirectories + 1
            ' Cast the object to a DirectoryInfo object.
            Dim dInfo As DotNetLib.DirectoryInfo
            Set dInfo = fileSytemInfoItem
            ' Iterate through all sub-directories.
            Call ListDirectoriesAndFiles(dInfo.GetFileSystemInfos)
            
        '/ Check to see if this is a FileInfo object.
        ElseIf TypeOf fileSytemInfoItem Is DotNetLib.FileInfo Then
            ' Add one to the file count.
            numFiles = numFiles + 1
        End If
    Next
End Sub
