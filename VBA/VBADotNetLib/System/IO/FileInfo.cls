VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides properties and instance methods for the creation, copying, deletion, moving, and opening of files, and aids in the creation of FileStream objects. This class cannot be inherited."
'@ModuleDescription("Provides properties and instance methods for the creation, copying, deletion, moving, and opening of files, and aids in the creation of FileStream objects. This class cannot be inherited.")
'@Folder "VBADotNetLib.System.IO"
'@PredeclaredId
'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 6, 2023
'@LastModified November 6, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.fileinfo?view=netframework-4.8.1

'@Remarks
'Singleton class for FileInfo containing factory methods and static members.

Option Explicit

Private Type TFileInfoSingleton
   FileInfoSingleton As DotNetLib.FileInfoSingleton
End Type

Private this As TFileInfoSingleton

Private Sub Class_Initialize()
   Set this.FileInfoSingleton = New DotNetLib.FileInfoSingleton
End Sub

Private Sub Class_Terminate()
   Set this.FileInfoSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the FileInfo class, which acts as a wrapper for a file path.")
Public Function Create(ByVal fileName As String) As DotNetLib.FileInfo
Attribute Create.VB_Description = "Initializes a new instance of the FileInfo class, which acts as a wrapper for a file path."
   Set Create = this.FileInfoSingleton.Create(fileName)
End Function
