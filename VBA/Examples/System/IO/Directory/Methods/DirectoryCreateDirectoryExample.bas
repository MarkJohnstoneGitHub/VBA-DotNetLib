Attribute VB_Name = "DirectoryCreateDirectoryExample"
'@Folder("Examples.System.IO.Directory.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 7, 2023
'@LastModified November 7, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.createdirectory?view=netframework-4.8.1#system-io-directory-createdirectory(system-string)

Option Explicit

''
' The following example creates and deletes the specified directory:
''
Public Sub DirectoryCreateDirectoryExample()
    ' Specify the directory you want to manipulate.
    Dim pvtPath As String
    pvtPath = "c:\MyDir"
    On Error Resume Next
    ' Determine whether the directory exists.
    If (Directory.Exists(pvtPath)) Then
        Debug.Print "That path exists already."
        GoTo CleanExit
    End If
    
    ' Try to create the directory.
    Dim di As DotNetLib.DirectoryInfo
    Set di = Directory.CreateDirectory(pvtPath)
    If Err.number = 0 Then
        Debug.Print VBAString.Format("The directory was created successfully at {0}.", Directory.GetCreationTime(pvtPath))

        ' Delete the directory.
        Call di.Delete
        Debug.Print "The directory was deleted successfully."
    Else
        Debug.Print VBAString.Format("The process failed: {0}", Err.Description)
    End If
CleanExit:
    On Error GoTo 0 'Stop code and display error
End Sub

