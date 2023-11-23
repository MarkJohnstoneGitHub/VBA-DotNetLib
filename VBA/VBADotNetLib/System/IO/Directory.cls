VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Directory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Exposes static methods for creating, moving, and enumerating through directories and subdirectories. This class cannot be inherited."
'@ModuleDescription("Exposes static methods for creating, moving, and enumerating through directories and subdirectories. This class cannot be inherited.")
'@Folder "VBADotNetLib.System.IO"
'@PredeclaredId
'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 6, 2023
'@LastModified November 17, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory?view=netframework-4.8.1

'@Remarks
'Singleton class for Directory containing factory methods and static members.

Option Explicit

Private Type TDirectorySingleton
   DirectorySingleton As DotNetLib.DirectorySingleton
End Type

Private this As TDirectorySingleton

Private Sub Class_Initialize()
   Set this.DirectorySingleton = New DotNetLib.DirectorySingleton
End Sub

Private Sub Class_Terminate()
   Set this.DirectorySingleton = Nothing
End Sub

'@Description("Creates all directories and subdirectories in the specified path unless they already exist.")
Public Function CreateDirectory(ByVal pPath As String) As DotNetLib.DirectoryInfo
Attribute CreateDirectory.VB_Description = "Creates all directories and subdirectories in the specified path unless they already exist."
   Set CreateDirectory = this.DirectorySingleton.CreateDirectory(pPath)
End Function

'@Description("Creates all the directories in the specified path, unless they already exist, applying the specified Windows security.")
Public Function CreateDirectory2(ByVal pPath As String, ByVal directorySecurity As Object) As DotNetLib.DirectoryInfo
Attribute CreateDirectory2.VB_Description = "Creates all the directories in the specified path, unless they already exist, applying the specified Windows security."
   Set CreateDirectory2 = this.DirectorySingleton.CreateDirectory_2(pPath, directorySecurity)
End Function

'@Description("Deletes an empty directory from a specified path.")
Public Sub Delete(ByVal pPath As String)
Attribute Delete.VB_Description = "Deletes an empty directory from a specified path."
    Call this.DirectorySingleton.Delete(pPath)
End Sub

'@Description("Deletes the specified directory and, if indicated, any subdirectories and files in the directory.")
Public Sub Delete2(ByVal pPath As String, ByVal recursive As Boolean)
Attribute Delete2.VB_Description = "Deletes the specified directory and, if indicated, any subdirectories and files in the directory."
    Call this.DirectorySingleton.Delete_2(pPath, recursive)
End Sub

'@Description("Returns an enumerable collection of directory full names that match a search pattern in a specified path, and optionally searches subdirectories.")
Public Function EnumerateDirectories(ByVal pPath As String, Optional ByVal searchPattern As String = "*", Optional ByVal pSearchOption As mscorlib.SearchOption = SearchOption.SearchOption_TopDirectoryOnly) As mscorlib.IEnumerable
Attribute EnumerateDirectories.VB_Description = "Returns an enumerable collection of directory full names that match a search pattern in a specified path, and optionally searches subdirectories."
   Set EnumerateDirectories = this.DirectorySingleton.EnumerateDirectories(pPath, searchPattern, pSearchOption)
End Function

'@Description("Returns an enumerable collection of full file names that match a search pattern in a specified path, and optionally searches subdirectories.")
Public Function EnumerateFiles(ByVal pPath As String, Optional ByVal searchPattern As String = "*", Optional ByVal pSearchOption As mscorlib.SearchOption = SearchOption.SearchOption_TopDirectoryOnly) As mscorlib.IEnumerable
Attribute EnumerateFiles.VB_Description = "Returns an enumerable collection of full file names that match a search pattern in a specified path, and optionally searches subdirectories."
   Set EnumerateFiles = this.DirectorySingleton.EnumerateFiles(pPath, searchPattern, pSearchOption)
End Function

'@Description("Returns an enumerable collection of file names and directory names that match a search pattern in a specified path, and optionally searches subdirectories.")
Public Function EnumerateFileSystemEntries(ByVal pPath As String, Optional ByVal searchPattern As String = "*", Optional ByVal pSearchOption As SearchOption = SearchOption.SearchOption_TopDirectoryOnly) As mscorlib.IEnumerable
Attribute EnumerateFileSystemEntries.VB_Description = "Returns an enumerable collection of file names and directory names that match a search pattern in a specified path, and optionally searches subdirectories."
   Set EnumerateFileSystemEntries = this.DirectorySingleton.EnumerateFileSystemEntries(pPath, searchPattern, pSearchOption)
End Function

'@Description("Determines whether the given path refers to an existing directory on disk.")
Public Function Exists(ByVal pPath As String) As Boolean
Attribute Exists.VB_Description = "Determines whether the given path refers to an existing directory on disk."
   Exists = this.DirectorySingleton.Exists(pPath)
End Function

'@Description("Gets a DirectorySecurity object that encapsulates the access control list (ACL) entries for a specified directory")
Public Function GetAccessControl(ByVal pPath As String) As Object
Attribute GetAccessControl.VB_Description = "Gets a DirectorySecurity object that encapsulates the access control list (ACL) entries for a specified directory"
   Set GetAccessControl = this.DirectorySingleton.GetAccessControl(pPath)
End Function

'@Description("Gets a DirectorySecurity object that encapsulates the specified type of access control list (ACL) entries for a specified directory.")
Public Function GetAccessControl2(ByVal pPath As String, ByVal includeSections As DotNetLib.AccessControlSections) As Object
Attribute GetAccessControl2.VB_Description = "Gets a DirectorySecurity object that encapsulates the specified type of access control list (ACL) entries for a specified directory."
   Set GetAccessControl2 = this.DirectorySingleton.GetAccessControl_2(pPath, includeSections)
End Function

'@Description("Gets the creation date and time of a directory.")
Public Function GetCreationTime(ByVal Path As String) As DotNetLib.DateTime
Attribute GetCreationTime.VB_Description = "Gets the creation date and time of a directory."
   Set GetCreationTime = this.DirectorySingleton.GetCreationTime(Path)
End Function

'@Description("Gets the creation date and time, in Coordinated Universal Time (UTC) format, of a directory.")
Public Function GetCreationTimeUtc(ByVal pPath As String) As DotNetLib.DateTime
Attribute GetCreationTimeUtc.VB_Description = "Gets the creation date and time, in Coordinated Universal Time (UTC) format, of a directory."
   Set GetCreationTimeUtc = this.DirectorySingleton.GetCreationTimeUtc(pPath)
End Function

'@Description("Gets the current working directory of the application.")
Public Function GetCurrentDirectory() As String
Attribute GetCurrentDirectory.VB_Description = "Gets the current working directory of the application."
   GetCurrentDirectory = this.DirectorySingleton.GetCurrentDirectory()
End Function

'@Description("Returns the names of the subdirectories (including their paths) that match the specified search pattern in the specified directory, and optionally searches subdirectories.")
Public Function GetDirectories(ByVal Path As String, Optional ByVal searchPattern As String = "*", Optional ByVal pSearchOption As mscorlib.SearchOption = mscorlib.SearchOption.SearchOption_TopDirectoryOnly) As String()
Attribute GetDirectories.VB_Description = "Returns the names of the subdirectories (including their paths) that match the specified search pattern in the specified directory, and optionally searches subdirectories."
   GetDirectories = this.DirectorySingleton.GetDirectories(Path, searchPattern, pSearchOption)
End Function

'@Description("Returns the volume information, root information, or both for the specified path.")
Public Function GetDirectoryRoot(ByVal pPath As String) As String
Attribute GetDirectoryRoot.VB_Description = "Returns the volume information, root information, or both for the specified path."
   GetDirectoryRoot = this.DirectorySingleton.GetDirectoryRoot(pPath)
End Function

'@Description("Returns the names of files (including their paths) that match the specified search pattern in the specified directory, using a value to determine whether to search subdirectories.")
Public Function GetFiles(ByVal pPath As String, Optional ByVal searchPattern As String = "*", Optional ByVal pSearchOption As SearchOption = mscorlib.SearchOption.SearchOption_TopDirectoryOnly) As String()
Attribute GetFiles.VB_Description = "Returns the names of files (including their paths) that match the specified search pattern in the specified directory, using a value to determine whether to search subdirectories."
   GetFiles = this.DirectorySingleton.GetFiles(pPath, searchPattern, pSearchOption)
End Function

'@Description("Returns an array of all the file names and directory names that match a search pattern in a specified path, and optionally searches subdirectories.")
Public Function GetFileSystemEntries(ByVal pPath As String, Optional ByVal searchPattern As String = "*", Optional ByVal pSearchOption As SearchOption = mscorlib.SearchOption.SearchOption_TopDirectoryOnly) As String()
Attribute GetFileSystemEntries.VB_Description = "Returns an array of all the file names and directory names that match a search pattern in a specified path, and optionally searches subdirectories."
   GetFileSystemEntries = this.DirectorySingleton.GetFileSystemEntries(pPath, searchPattern, pSearchOption)
End Function

'@Description("Returns the date and time the specified file or directory was last accessed.")
Public Function GetLastAccessTime(ByVal pPath As String) As DotNetLib.DateTime
Attribute GetLastAccessTime.VB_Description = "Returns the date and time the specified file or directory was last accessed."
   Set GetLastAccessTime = this.DirectorySingleton.GetLastAccessTime(pPath)
End Function

'@Description("Returns the date and time, in Coordinated Universal Time (UTC) format, that the specified file or directory was last accessed.")
Public Function GetLastAccessTimeUtc(ByVal pPath As String) As DotNetLib.DateTime
Attribute GetLastAccessTimeUtc.VB_Description = "Returns the date and time, in Coordinated Universal Time (UTC) format, that the specified file or directory was last accessed."
   Set GetLastAccessTimeUtc = this.DirectorySingleton.GetLastAccessTimeUtc(pPath)
End Function

'@Description("Returns the date and time the specified file or directory was last written to.")
Public Function GetLastWriteTime(ByVal pPath As String) As DotNetLib.DateTime
Attribute GetLastWriteTime.VB_Description = "Returns the date and time the specified file or directory was last written to."
   Set GetLastWriteTime = this.DirectorySingleton.GetLastWriteTime(pPath)
End Function

'@Description("Returns the date and time, in Coordinated Universal Time (UTC) format, that the specified file or directory was last written to.")
Public Function GetLastWriteTimeUtc(ByVal pPath As String) As DotNetLib.DateTime
Attribute GetLastWriteTimeUtc.VB_Description = "Returns the date and time, in Coordinated Universal Time (UTC) format, that the specified file or directory was last written to."
   Set GetLastWriteTimeUtc = this.DirectorySingleton.GetLastWriteTimeUtc(pPath)
End Function

'@Description("Retrieves the names of the logical drives on this computer in the form "<drive letter>:\".")
Public Function GetLogicalDrives() As String()
   GetLogicalDrives = this.DirectorySingleton.GetLogicalDrives()
End Function

'@Description("Retrieves the parent directory of the specified path, including both absolute and relative paths.")
Public Function GetParent(ByVal pPath As String) As DotNetLib.DirectoryInfo
Attribute GetParent.VB_Description = "Retrieves the parent directory of the specified path, including both absolute and relative paths."
   Set GetParent = this.DirectorySingleton.GetParent(pPath)
End Function

'@Description("Moves a file or a directory and its contents to a new location.")
Public Sub Move(ByVal sourceDirName As String, ByVal destDirName As String)
Attribute Move.VB_Description = "Moves a file or a directory and its contents to a new location."
    Call this.DirectorySingleton.Move(sourceDirName, destDirName)
End Sub

'@Description("Applies access control list (ACL) entries described by a DirectorySecurity object to the specified directory.")
Public Sub SetAccessControl(ByVal pPath As String, ByVal directorySecurity As Object)
Attribute SetAccessControl.VB_Description = "Applies access control list (ACL) entries described by a DirectorySecurity object to the specified directory."
    Call this.DirectorySingleton.SetAccessControl(pPath, directorySecurity)
End Sub

'@Description("Sets the creation date and time for the specified file or directory.")
Public Sub SetCreationTime(ByVal pPath As String, ByVal creationTime As DotNetLib.DateTime)
Attribute SetCreationTime.VB_Description = "Sets the creation date and time for the specified file or directory."
    Call this.DirectorySingleton.SetCreationTime(pPath, creationTime)
End Sub

'@Description("Sets the creation date and time, in Coordinated Universal Time (UTC) format, for the specified file or directory.")
Public Sub SetCreationTimeUtc(ByVal pPath As String, ByVal creationTimeUtc As DotNetLib.DateTime)
Attribute SetCreationTimeUtc.VB_Description = "Sets the creation date and time, in Coordinated Universal Time (UTC) format, for the specified file or directory."
    Call this.DirectorySingleton.SetCreationTimeUtc(pPath, creationTimeUtc)
End Sub

'@Description("Sets the application's current working directory to the specified directory.")
Public Sub SetCurrentDirectory(ByVal Path As String)
Attribute SetCurrentDirectory.VB_Description = "Sets the application's current working directory to the specified directory."
    Call this.DirectorySingleton.SetCurrentDirectory(Path)
End Sub

'@Description("Sets the date and time the specified file or directory was last accessed.")
Public Sub SetLastAccessTime(ByVal pPath As String, ByVal lastAccessTime As DotNetLib.DateTime)
Attribute SetLastAccessTime.VB_Description = "Sets the date and time the specified file or directory was last accessed."
    Call this.DirectorySingleton.SetLastAccessTime(pPath, lastAccessTime)
End Sub

'@Description("Sets the date and time, in Coordinated Universal Time (UTC) format, that the specified file or directory was last accessed.")
Public Sub SetLastAccessTimeUtc(ByVal pPath As String, ByVal lastAccessTimeUtc As DotNetLib.DateTime)
Attribute SetLastAccessTimeUtc.VB_Description = "Sets the date and time, in Coordinated Universal Time (UTC) format, that the specified file or directory was last accessed."
    Call this.DirectorySingleton.SetLastAccessTimeUtc(pPath, lastAccessTimeUtc)
End Sub

'@Description("Sets the date and time a directory was last written to.")
Public Sub SetLastWriteTime(ByVal pPath As String, ByVal lastWriteTime As DotNetLib.DateTime)
Attribute SetLastWriteTime.VB_Description = "Sets the date and time a directory was last written to."
    Call this.DirectorySingleton.SetLastWriteTime(pPath, lastWriteTime)
End Sub

'@Description("Sets the date and time, in Coordinated Universal Time (UTC) format, that a directory was last written to.")
Public Sub SetLastWriteTimeUtc(ByVal pPath As String, ByVal lastWriteTimeUtc As DotNetLib.DateTime)
Attribute SetLastWriteTimeUtc.VB_Description = "Sets the date and time, in Coordinated Universal Time (UTC) format, that a directory was last written to."
    Call this.DirectorySingleton.SetLastWriteTimeUtc(pPath, lastWriteTimeUtc)
End Sub
