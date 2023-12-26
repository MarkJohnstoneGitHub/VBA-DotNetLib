Attribute VB_Name = "DirInfoDeleteExample"
'@Folder("Examples.System.IO.DirectoryInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 23, 2023
'@LastModified December 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.delete?view=netframework-4.8.1#system-io-directoryinfo-delete

Option Explicit

''
' The following example throws an exception if you attempt to delete a directory
' that is not empty.
''
Public Sub DirectoryInfoDeleteExample()
    ' Specify the directories you want to manipulate.
    Dim di1 As DotNetLib.DirectoryInfo
    Set di1 = DirectoryInfo.Create("c:\MyDir")
    
    On Error GoTo ErrorHandler
    ' Create the directories.
    Call di1.Create
    Call di1.CreateSubdirectory("temp")

    'This operation will not be allowed because there are subdirectories.
    Debug.Print VBAString.Format("I am about to attempt to delete {0}", di1.name)
    Call di1.Delete
    Debug.Print "The Delete operation was successful, which was unexpected."
Exit Sub
ErrorHandler:
    Debug.Print "The Delete operation failed as expected."
End Sub
