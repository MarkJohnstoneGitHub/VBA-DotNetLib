VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DirectoryInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Exposes instance methods for creating, moving, and enumerating through directories and subdirectories. This class cannot be inherited."
'@ModuleDescription("Exposes instance methods for creating, moving, and enumerating through directories and subdirectories. This class cannot be inherited.")
'@Folder "VBADotNetLib.System.IO"
'@PredeclaredId
'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 6, 2023
'@LastModified November 6, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo?view=netframework-4.8.1

'@Remarks
'Singleton class for DirectoryInfo containing factory methods and static members.

Option Explicit

Private Type TDirectoryInfoSingleton
   DirectoryInfoSingleton As DotNetLib.DirectoryInfoSingleton
End Type

Private this As TDirectoryInfoSingleton

Private Sub Class_Initialize()
   Set this.DirectoryInfoSingleton = New DotNetLib.DirectoryInfoSingleton
End Sub

Private Sub Class_Terminate()
   Set this.DirectoryInfoSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the DirectoryInfo class on the specified path.")
Public Function Create(ByVal Path As String) As DotNetLib.DirectoryInfo
Attribute Create.VB_Description = "Initializes a new instance of the DirectoryInfo class on the specified path."
   Set Create = this.DirectoryInfoSingleton.Create(Path)
End Function
