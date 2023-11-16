Attribute VB_Name = "DirectoryEnumerateDirectories1"
'@Folder("Examples.System.IO.Directory.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 8, 2023
'@LastModified November 9, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.enumeratedirectories?view=netframework-4.8.1#system-io-directory-enumeratedirectories(system-string)

Option Explicit

''
'  The following example enumerates the top-level directories in a specified path.
''
Public Sub DirectoryEnumerateDirectories()
    On Error GoTo ErrorHandler
    ' Set a variable to the My Documents path.
    Dim docPath As String
    docPath = Environment.GetFolderPath(SpecialFolder.SpecialFolder_MyDocuments)
    Dim dirs As DotNetLib.ListString
    Set dirs = ListString.CreateFromIEnumerable(Directory.EnumerateDirectories(docPath))
    
    Dim dirSeparator As DotNetLib.String
    Set dirSeparator = Strings.Create(Path.DirectorySeparatorChar)
    Debug.Print "Document Path: "; docPath
    Dim varDir As Variant
    For Each varDir In dirs
        Dim dir As DotNetLib.String
        Set dir = Strings.Create(varDir)
        Debug.Print VBAString.Format("{0}", dir.Substring(dir.LastIndexOf(dirSeparator) + 1))
    Next
    Debug.Print VBAString.Format("{0} directories found.", dirs.count);
    Exit Sub
ErrorHandler:
    If Err.number = COMHResult.UnauthorizedAccessException Then
        Debug.Print Err.Description
    End If
    If Err.number = COMHResult.PathTooLongException Then
        Debug.Print Err.Description
    End If
End Sub
