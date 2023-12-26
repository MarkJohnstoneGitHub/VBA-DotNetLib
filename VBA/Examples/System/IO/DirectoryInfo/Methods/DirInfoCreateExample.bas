Attribute VB_Name = "DirInfoCreateExample"
'@Folder("Examples.System.IO.DirectoryInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 23, 2023
'@LastModified December 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.create?view=netframework-4.8.1#system-io-directoryinfo-create

Option Explicit

''
' The following example checks whether a specified directory exists, creates
' the directory if it does not exist, and deletes the directory.
''
Public Sub DirectoryInfoCreateExample()
    ' Specify the directories you want to manipulate.
    Dim di As DotNetLib.DirectoryInfo
    Set di = DirectoryInfo.Create("c:\MyDir")
    
    On Error GoTo ErrorHandler
    '/ Determine whether the directory exists.
    If (di.Exists) Then
        ' Indicate that it already exists.
        Debug.Print "That path exists already."
        Exit Sub 'Return
    End If
    
    ' Try to create the directory.
    Call di.Create
    Debug.Print "The directory was created successfully."
    
    ' Delete the directory.
    Call di.Delete
    Debug.Print "The directory was deleted successfully."
Exit Sub
ErrorHandler:
    Debug.Print "The process failed: {0}", Err.Description
End Sub
