Attribute VB_Name = "DirectoryEnumerateFilesEg2"
'@Folder("Examples.System.IO.Directory.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 14, 2023
'@LastModified November 14, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference

Option Explicit

''
' The following example enumerates the files in the specified directory,
' reads each line of the file, and displays the line if it contains the
' string "Europe".
''

Public Sub DirectoryEnumerateFilesExample2()
    On Error GoTo ErrorHandler
    Dim txtFiles As mscorlib.IEnumerable
    Set txtFiles = Directory.EnumerateFiles("C:\VBA\Export")
    
    Dim files As DotNetLib.ListString
    Set files = ListString.Create()
'    Set files = ListString.CreateFromIEnumerable(txtFiles)

    Dim varCurrentFile As Variant
    For Each varCurrentFile In txtFiles
        Dim currentFile As DotNetLib.String
        Set currentFile = Strings.Create(varCurrentFile)
        Dim fileName As String
        Dim txtFile() As String
        txtFile = File.ReadAllLines(varCurrentFile)
        Dim varLine As Variant
        For Each varLine In txtFile
            Dim line As DotNetLib.String
            Set line = Strings.Create(varLine)
            If line.ToLower().Contains_2("on error") Then
                Debug.Print "Contains On Error"
            End If
        Next
        
        
'            // LINQ query for all files containing the word 'Europe'.
'            var files = from file in
'                Directory.EnumerateFiles(@"\\archives1\library\")
'                where File.ToLower().Contains("europe")
'                select file;


        'Function ReadAllLines(ByVal pPath As String) As String()
'        fileName = currentFile.Substring(sourceDirectory.length + 1).ToString
'        Call Directory.Move(currentFile.ToString, Path.Combine2(archiveDirectory, fileName))
    Next
    Exit Sub
ErrorHandler:
    Debug.Print Err.Description
End Sub

'    On Error GoTo ErrorHandler
'    Dim dirPath As String
'    dirPath = "\\archives\2009\reports"
'    Dim dirs As DotNetLib.ListString
'    Set dirs = ListString.CreateFromIEnumerable( _
'                Directory.EnumerateDirectories(dirPath, "dv_*", _
'                                                SearchOption.SearchOption_AllDirectories))
'    ' Show results.
'    Dim varDir As Variant
'    For Each varDir In dirs
'        Dim dir As DotNetLib.String
'        Set dir = Strings.Create(varDir)
'        ' // Remove path information from string.
'        Debug.Print BString.Format("{0}", _
'                             dir.Substring(dir.LastIndexOf_7("\") + 1))
'    Next
'    Debug.Print BString.Format("{0} directories found.", dirs.Count);
'
'    Exit Sub
'ErrorHandler:
'    Debug.Print Err.Description



'using System;
'using System.Collections.Generic;
'using System.Linq;
'using System.IO;
'
'Class Program
'{
'    static void Main(string[] args)
'    {
'        Try
'        {
'            // LINQ query for all files containing the word 'Europe'.
'            var files = from file in
'                Directory.EnumerateFiles(@"\\archives1\library\")
'                where File.ToLower().Contains("europe")
'                select file;
'
'            foreach (var file in files)
'            {
'                Console.WriteLine("{0}", file);
'            }
'            Console.WriteLine("{0} files found.", files.Count<string>().ToString());
'        }
'        catch (UnauthorizedAccessException UAEx)
'        {
'            Console.WriteLine(UAEx.Message);
'        }
'        catch (PathTooLongException PathEx)
'        {
'            Console.WriteLine(PathEx.Message);
'        }
'    }
'}
