Attribute VB_Name = "DirectoryExistsExample"
'@Folder("Examples.System.IO.Directory.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 14, 2023
'@LastModified November 14, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.exists?view=netframework-4.8.1#examples

Option Explicit

''
' The following example takes an array of file or directory names on the command
' line, determines what kind of name it is, and processes it appropriately.
''
Public Sub DirectoryExistsExample()
    Dim args() As String
    Call ArrayEx.CreateInitialize1D(args, "C:\Current")
    Dim varPath As Variant
    For Each varPath In args
        If File.Exists(varPath) Then
            ' This path is a file
            Call ProcessFile(varPath)
        ElseIf (Directory.Exists(varPath)) Then
            Call ProcessDirectory(varPath)
        Else
            Debug.Print VBAString.Format("{0} is not a valid file or directory.", varPath)
        End If
    Next
End Sub

' Process all files in the directory passed in, recurse on any directories
' that are found, and process the files they contain.
Public Sub ProcessDirectory(ByVal targetDirectory As String)
    Dim fileEntries() As String
    fileEntries = Directory.GetFiles(targetDirectory)
    Dim varFileName As Variant
    For Each varFileName In fileEntries
        Call ProcessFile(varFileName)
    Next
    
    ' Recurse into subdirectories of this directory.
    Dim subdirectoryEntries() As String
    subdirectoryEntries = Directory.GetDirectories(targetDirectory)
    Dim varSubdirectory As Variant
    For Each varSubdirectory In subdirectoryEntries
        Call ProcessDirectory(varSubdirectory)
    Next
End Sub

' Insert logic for processing found files here.
Public Sub ProcessFile(ByVal pPath As String)
    Debug.Print VBAString.Format("Processed file '{0}'.", pPath)
End Sub
