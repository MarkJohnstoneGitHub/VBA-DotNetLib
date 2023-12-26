Attribute VB_Name = "DirectoryEnumerateDirectories3"
'@Folder("Examples.System.IO.Directory.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 12, 2023
'@LastModified November 16, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.enumeratedirectories?view=netframework-4.8.1#system-io-directory-enumeratedirectories(system-string-system-string-system-io-searchoption)

Option Explicit

Public Sub DirectoryEnumerateDirectoriesEg3()
    On Error GoTo ErrorHandler
    Dim dirPath As String
    dirPath = "C:\VBA\Output\Examples" 'Eg Directory containing the exported project according to Rubberduck folder annotations
    Dim dirs As DotNetLib.ListString
    
     'Obtain directories and subdirectories beginning with "Date"
    Set dirs = ListString.CreateFromIEnumerable( _
                Directory.EnumerateDirectories(dirPath, "Date*", _
                                                SearchOption.SearchOption_AllDirectories))
                                                
    ' Show results.
    Dim varDir As Variant
    For Each varDir In dirs
        Dim dir As DotNetLib.String
        Set dir = Strings.Create(varDir)
        ' // Remove path information from string.
        Debug.Print VBAString.Format("{0}", _
                             dir.Substring(dir.LastIndexOf_7("\") + 1))
    Next
    Debug.Print VBAString.Format("{0} directories found.", dirs.Count);
    
    Exit Sub
ErrorHandler:
    Debug.Print Err.Description
End Sub

' Example Output:
'   DateTime
'   DateTimeOffset
'   DateTimeFormatInfo
'   3 directories found.

