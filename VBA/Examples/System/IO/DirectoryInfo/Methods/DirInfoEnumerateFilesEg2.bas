Attribute VB_Name = "DirInfoEnumerateFilesEg2"
'@Folder("Examples.System.IO.DirectoryInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 26, 2023
'@LastModified December 26, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.enumeratefiles?view=netframework-4.8.1#system-io-directoryinfo-enumeratefiles(system-string)

Option Explicit

''
' The following example shows how to enumerate files in a directory by using
' different search options. The example assumes a directory that has files
' named log1.txt, log2.txt, test1.txt, test2.txt, test3.txt, and a subdirectory
' that has a file named SubFile.txt.
''
Public Sub DirectoryInfoEnumerateFilesExample2()
    Dim di As DotNetLib.DirectoryInfo
    Set di = DirectoryInfo.Create("C:\ExampleDir")
    Debug.Print "No search pattern returns:"
    
    Dim varFileInfo As Variant
    For Each varFileInfo In di.EnumerateFiles()
        Dim fi As DotNetLib.FileInfo
        Set fi = varFileInfo
        Debug.Print fi.name
    Next
    Debug.Print
    
    Debug.Print "Search pattern *2* returns:"
    For Each varFileInfo In di.EnumerateFiles("*2*")
        Set fi = varFileInfo
        Debug.Print fi.name
    Next
    Debug.Print

    Debug.Print "Search pattern test?.txt returns:"
    For Each varFileInfo In di.EnumerateFiles("test?.txt")
        Set fi = varFileInfo
        Debug.Print fi.name
    Next
    Debug.Print

    Debug.Print "Search pattern AllDirectories returns:"
    For Each varFileInfo In di.EnumerateFiles("*", SearchOption.SearchOption_AllDirectories)
        Set fi = varFileInfo
        Debug.Print fi.name
    Next
End Sub

'/*
'This code produces output similar to the following:
'
'No search pattern returns:
'log1.txt
'log2.txt
'test1.txt
'test2.txt
'test3.txt
'
'Search pattern * 2 * returns:
'log2.txt
'test2.txt
'
'Search pattern test?.txt returns:
'test1.txt
'test2.txt
'test3.txt
'
'Search pattern AllDirectories returns:
'log1.txt
'log2.txt
'test1.txt
'test2.txt
'test3.txt
'SubFile.txt
'Press any key to continue . . .
'
'*/
