Attribute VB_Name = "DirectoryGetFileSystemEntriesEg"
'@Folder "Examples.System.IO.Directory.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 18, 2023
'@LastModified January 28, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.getfilesystementries?view=netframework-4.8.1#system-io-directory-getfilesystementries(system-string)

Option Explicit

''
' Returns the names of all files and subdirectories in a specified path.
''
Public Sub DirectoryGetFileSystemEntries()
    Dim pvtPath As String
    pvtPath = Directory.GetCurrentDirectory()
    Dim filter As String
    filter = "*.exe"
    Call PrintFileSystemEntries(pvtPath)
    Call PrintFileSystemEntries2(pvtPath, filter)
    Call GetLogicalDrives
    Call GetParent(pvtPath)
    Call Move("C:\\proof", "C:\\Temp")
End Sub

Private Sub PrintFileSystemEntries(ByVal pPath As String)
    On Error GoTo ErrorHandler
    ' Obtain the file system entries in the directory path.
    Dim directoryEntries() As String
    directoryEntries = Directory.GetFileSystemEntries(pPath)
    
    Dim str As Variant
    For Each str In directoryEntries
        Debug.Print str
    Next
Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case ArgumentNullException
            Debug.Print "Path is a null reference."
        Case SecurityException
            Debug.Print "The caller does not have the required permission."
        Case ArgumentException
            Debug.Print "Path is an empty string, " + _
                        "contains only white spaces, " + _
                        "or contains invalid characters."
        Case DirectoryNotFoundException
            Debug.Print "The path encapsulated in the " + _
                        "Directory object does not exist."
    End Select
End Sub

Private Sub PrintFileSystemEntries2(ByVal pPath As String, ByVal pattern As String)
    On Error GoTo ErrorHandler
    ' Obtain the file system entries in the directory path.
    Dim directoryEntries() As String
    directoryEntries = Directory.GetFileSystemEntries(pPath, pattern)
    Dim str As Variant
    For Each str In directoryEntries
        Debug.Print str
    Next
Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case ArgumentNullException
            Debug.Print "Path is a null reference."
        Case SecurityException
            Debug.Print "The caller does not have the required permission."
        Case ArgumentException
            Debug.Print "Path is an empty string, " + _
                        "contains only white spaces, " + _
                        "or contains invalid characters."
        Case DirectoryNotFoundException
            Debug.Print "The path encapsulated in the " + _
                        "Directory object does not exist."
        Case Else
            Debug.Print Err.Description
    End Select
End Sub

' Print out all logical drives on the system.
Private Sub GetLogicalDrives()
    On Error GoTo ErrorHandler
    Dim drives() As String
    drives = Directory.GetLogicalDrives()
    Dim str As Variant
    For Each str In drives
        Debug.Print str
    Next
Exit Sub

ErrorHandler:
    If Err.Number = IOException Then
        Debug.Print "An I/O error occurs."
    ElseIf Err.Number = SecurityException Then
        Debug.Print "The caller does not have the required permission."
    End If
End Sub

Private Sub GetParent(ByVal pPath As String)
    On Error GoTo ErrorHandler
    Dim pvtDirectoryInfo As DotNetLib.DirectoryInfo
    Set pvtDirectoryInfo = Directory.GetParent(pPath)
    Debug.Print pvtDirectoryInfo.FullName
Exit Sub
ErrorHandler:
    If Err.Number = ArgumentNullException Then
        Debug.Print "Path is a null reference."
    ElseIf Err.Number = ArgumentException Then
        Debug.Print "Path is an empty string, " + _
                    "contains only white spaces, or " + _
                    "contains invalid characters."
    End If
End Sub

Private Sub Move(ByVal sourcePath As String, ByVal destinationPath As String)
    On Error GoTo ErrorHandler
    Call Directory.Move(sourcePath, destinationPath)
    Debug.Print "The directory move is complete."
Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case ArgumentNullException
            Debug.Print "Path is a null reference."
        Case SecurityException
            Debug.Print "The caller does not have the required permission."
        Case ArgumentException
            Debug.Print "Path is an empty string, " + _
                        "contains only white spaces, " + _
                        "or contains invalid characters."
        Case IOException
            Debug.Print "An attempt was made to move a " + _
                        "directory to a different " + _
                        "volume, or destDirName " + _
                        "already exists."
        Case Else
            Debug.Print Err.Description
    End Select
End Sub


