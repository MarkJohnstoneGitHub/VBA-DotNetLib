Attribute VB_Name = "DirInfoConstructorExample"
'@Folder("Examples.System.IO.DirectoryInfo.Constructors")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 22, 2023
'@LastModified December 22, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.-ctor?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses this constructor to create the specified directory
' and subdirectory, and demonstrates that a directory that contains
' subdirectories cannot be deleted.
''
Public Sub DirectoryInfoConstructor()
    ' Specify the directories you want to manipulate.
    Dim di1 As DotNetLib.DirectoryInfo
    Set di1 = DirectoryInfo.Create("c:\MyDir")
    Dim di2 As DotNetLib.DirectoryInfo
    Set di2 = DirectoryInfo.Create("c:\MyDir\temp")
    On Error GoTo ErrorHandler
    ' Create the directories.
    Call di1.Create
    Call di2.Create
    
    ' This operation will not be allowed because there are subdirectories.
    Debug.Print VBAString.Format("I am about to attempt to delete {0}.", di1.Name)
    Call di1.Delete
    Debug.Print "The Delete operation was successful, which was unexpected."
    
    Exit Sub
ErrorHandler:
    Debug.Print "The Delete operation failed as expected."
End Sub
