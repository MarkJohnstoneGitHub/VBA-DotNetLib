Attribute VB_Name = "DirInfoGetFilesExample"
'@Folder "Examples.System.IO.DirectoryInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 26, 2023
'@LastModified December 26, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.getfiles?view=netframework-4.8.1#system-io-directoryinfo-getfiles

Option Explicit

''
' The following example shows how to get a list of files from a directory by
' using different search options. The example assumes a directory that has
' files named log1.txt, log2.txt, test1.txt, test2.txt, test3.txt, and a
' subdirectory that has a file named SubFile.txt.
''
Public Sub DirectoryInfoGetFilesExample()
    Dim di As DotNetLib.DirectoryInfo
    Set di = DirectoryInfo.Create("C:\Users\tomfitz\Documents\ExampleDir")
    Debug.Print "No search pattern returns:"
    Dim varfi As Variant
    Dim fi As DotNetLib.FileInfo
    
    On Error Resume Next
    For Each varfi In di.GetFiles()
        Set fi = varfi
        Debug.Print fi.name
    Next

    Debug.Print
    
    Debug.Print "Search pattern *2* returns:"
    For Each varfi In di.GetFiles("*2*")
        Set fi = varfi
        Debug.Print fi.name
    Next
    
    Debug.Print
    Debug.Print "Search pattern AllDirectories returns:"
    For Each varfi In di.GetFiles("*", SearchOption.SearchOption_AllDirectories)
        Debug.Print fi.name
    Next
    On Error GoTo 0 'Stop code and display error
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
