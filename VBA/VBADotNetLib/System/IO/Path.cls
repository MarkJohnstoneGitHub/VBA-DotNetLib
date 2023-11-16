VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Path"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Performs operations on String instances that contain file or directory path information. These operations are performed in a cross-platform manner."
'@ModuleDescription("Performs operations on String instances that contain file or directory path information. These operations are performed in a cross-platform manner.")
'@Folder "VBADotNetLib.System.IO"
'@PredeclaredId
'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 6, 2023
'@LastModified November 6, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path?view=netframework-4.8.1

'@Remarks
'Singleton class for Path containing factory methods and static members.

Option Explicit

Private Type TPathSingleton
   PathSingleton As DotNetLib.PathSingleton
End Type

Private this As TPathSingleton

Private Sub Class_Initialize()
   Set this.PathSingleton = New DotNetLib.PathSingleton
End Sub

Private Sub Class_Terminate()
   Set this.PathSingleton = Nothing
End Sub

'@Description("Provides a platform-specific alternate character used to separate directory levels in a path string that reflects a hierarchical file system organization.")
Public Property Get AltDirectorySeparatorChar() As String
Attribute AltDirectorySeparatorChar.VB_Description = "Provides a platform-specific alternate character used to separate directory levels in a path string that reflects a hierarchical file system organization."
   AltDirectorySeparatorChar = this.PathSingleton.AltDirectorySeparatorChar()
End Property

'@Description("Provides a platform-specific character used to separate directory levels in a path string that reflects a hierarchical file system organization.")
Public Property Get DirectorySeparatorChar() As String
Attribute DirectorySeparatorChar.VB_Description = "Provides a platform-specific character used to separate directory levels in a path string that reflects a hierarchical file system organization."
   DirectorySeparatorChar = this.PathSingleton.DirectorySeparatorChar()
End Property

'@Description("A platform-specific separator character used to separate path strings in environment variables.")
Public Property Get PathSeparator() As String
Attribute PathSeparator.VB_Description = "A platform-specific separator character used to separate path strings in environment variables."
   PathSeparator = this.PathSingleton.PathSeparator()
End Property

'@Description("Provides a platform-specific volume separator character.")
Public Property Get VolumeSeparatorChar() As String
Attribute VolumeSeparatorChar.VB_Description = "Provides a platform-specific volume separator character."
   VolumeSeparatorChar = this.PathSingleton.VolumeSeparatorChar()
End Property

'@Description("Changes the extension of a path string.")
Public Function ChangeExtension(ByVal Path As String, ByVal extension As String) As String
Attribute ChangeExtension.VB_Description = "Changes the extension of a path string."
   ChangeExtension = this.PathSingleton.ChangeExtension(Path, extension)
End Function

'@Description("Combines an array of strings into a path.")
Public Function Combine(ByRef paths() As String) As String
Attribute Combine.VB_Description = "Combines an array of strings into a path."
   Combine = this.PathSingleton.Combine(paths)
End Function

'@Description("Combines two strings into a path.")
Public Function Combine2(ByVal path1 As String, ByVal path2 As String) As String
Attribute Combine2.VB_Description = "Combines two strings into a path."
   Combine2 = this.PathSingleton.Combine_2(path1, path2)
End Function

'@Description("Combines three strings into a path.")
Public Function Combine3(ByVal path1 As String, ByVal path2 As String, ByVal path3 As String) As String
Attribute Combine3.VB_Description = "Combines three strings into a path."
   Combine3 = this.PathSingleton.Combine_3(path1, path2, path3)
End Function

'@Description("Combines four strings into a path.")
Public Function Combine4(ByVal path1 As String, ByVal path2 As String, ByVal path3 As String, ByVal path4 As String) As String
Attribute Combine4.VB_Description = "Combines four strings into a path."
   Combine4 = this.PathSingleton.Combine_4(path1, path2, path3, path4)
End Function

'@Description("Returns the directory information for the specified path.")
Public Function GetDirectoryName(ByVal Path As String) As String
Attribute GetDirectoryName.VB_Description = "Returns the directory information for the specified path."
   GetDirectoryName = this.PathSingleton.GetDirectoryName(Path)
End Function

'@Description("Returns the extension (including the period .) of the specified path string.")
Public Function GetExtension(ByVal Path As String) As String
Attribute GetExtension.VB_Description = "Returns the extension (including the period .) of the specified path string."
   GetExtension = this.PathSingleton.GetExtension(Path)
End Function

'@Description("Returns the file name and extension of the specified path string.")
Public Function GetFileName(ByVal Path As String) As String
Attribute GetFileName.VB_Description = "Returns the file name and extension of the specified path string."
   GetFileName = this.PathSingleton.GetFileName(Path)
End Function

'@Description("Returns the file name of the specified path string without the extension.")
Public Function GetFileNameWithoutExtension(ByVal Path As String) As String
Attribute GetFileNameWithoutExtension.VB_Description = "Returns the file name of the specified path string without the extension."
   GetFileNameWithoutExtension = this.PathSingleton.GetFileNameWithoutExtension(Path)
End Function

'@Description("Returns the absolute path for the specified path string.")
Public Function GetFullPath(ByVal Path As String) As String
Attribute GetFullPath.VB_Description = "Returns the absolute path for the specified path string."
   GetFullPath = this.PathSingleton.GetFullPath(Path)
End Function

'@Description("Gets an array containing the characters that are not allowed in file names.")
Public Function GetInvalidFileNameChars() As String()
Attribute GetInvalidFileNameChars.VB_Description = "Gets an array containing the characters that are not allowed in file names."
   GetInvalidFileNameChars = this.PathSingleton.GetInvalidFileNameChars()
End Function

'@Description("Gets an array containing the characters that are not allowed in path names.")
Public Function GetInvalidPathChars() As String()
Attribute GetInvalidPathChars.VB_Description = "Gets an array containing the characters that are not allowed in path names."
   GetInvalidPathChars = this.PathSingleton.GetInvalidPathChars()
End Function

'@Description("Gets the root directory information from the path contained in the specified string.")
Public Function GetPathRoot(ByVal Path As String) As String
Attribute GetPathRoot.VB_Description = "Gets the root directory information from the path contained in the specified string."
   GetPathRoot = this.PathSingleton.GetPathRoot(Path)
End Function

'@Description("Returns a random folder name or file name.")
Public Function GetRandomFileName() As String
Attribute GetRandomFileName.VB_Description = "Returns a random folder name or file name."
   GetRandomFileName = this.PathSingleton.GetRandomFileName()
End Function

'@Description("Creates a uniquely named, zero-byte temporary file on disk and returns the full path of that file.")
Public Function GetTempFileName() As String
Attribute GetTempFileName.VB_Description = "Creates a uniquely named, zero-byte temporary file on disk and returns the full path of that file."
   GetTempFileName = this.PathSingleton.GetTempFileName()
End Function

'@Description("Returns the path of the current user's temporary folder.")
Public Function GetTempPath() As String
Attribute GetTempPath.VB_Description = "Returns the path of the current user's temporary folder."
   GetTempPath = this.PathSingleton.GetTempPath()
End Function

'@Description("Determines whether a path includes a file name extension.")
Public Function HasExtension(ByVal Path As String) As Boolean
Attribute HasExtension.VB_Description = "Determines whether a path includes a file name extension."
   HasExtension = this.PathSingleton.HasExtension(Path)
End Function

'@Description("Returns a value indicating whether the specified path string contains a root.")
Public Function IsPathRooted(ByVal Path As String) As Boolean
Attribute IsPathRooted.VB_Description = "Returns a value indicating whether the specified path string contains a root."
   IsPathRooted = this.PathSingleton.IsPathRooted(Path)
End Function
