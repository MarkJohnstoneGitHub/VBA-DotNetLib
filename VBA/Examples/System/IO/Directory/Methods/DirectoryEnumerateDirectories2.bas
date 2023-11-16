Attribute VB_Name = "DirectoryEnumerateDirectories2"
'@Folder("Examples.System.IO.Directory.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 12, 2023
'@LastModified November 12, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.enumeratedirectories?view=netframework-4.8.1#system-io-directory-enumeratedirectories(system-string-system-string)

Option Explicit

''
' The following example enumerates the top-level directories in a specified
' path that match a specified search pattern.
''
Public Sub DirectoryEnumerateDirectoriesEg2()
    On Error GoTo ErrorHandler
    Dim dirPath As String
    dirPath = "\\archives\2009\reports"
    Dim dirs As DotNetLib.ListString
    Set dirs = ListString.CreateFromIEnumerable(Directory.EnumerateDirectories(dirPath, "dv_*"))
    
    ' Show results.
    Dim varDir As Variant
    For Each varDir In dirs
        Dim dir As DotNetLib.String
        Set dir = Strings.Create(varDir)
        ' // Remove path information from string.
        Debug.Print VBAString.Format("{0}", _
                             dir.Substring(dir.LastIndexOf_7("\") + 1))
    Next
    Debug.Print VBAString.Format("{0} directories found.", dirs.count);
    
    Exit Sub
ErrorHandler:
    Debug.Print Err.Description
End Sub

'using System;
'using System.Collections.Generic;
'using System.IO;
'using System.Linq;
'
'Class Program
'{
'
'    private static void Main(string[] args)
'    {
'        Try
'        {
'            string dirPath = @"\\archives\2009\reports";
'
'            // LINQ query.
'            var dirs = from dir in
'                     Directory.EnumerateDirectories(dirPath, "dv_*")
'                       select dir;
'
'            // Show results.
'            foreach (var dir in dirs)
'            {
'                // Remove path information from string.
'                Console.WriteLine("{0}",
'                    dir.Substring(dir.LastIndexOf("\\") + 1));
'            }
'            Console.WriteLine("{0} directories found.",
'                dirs.Count<string>().ToString());
'
'            // Optionally create a List collection.
'            List<string> workDirs = new List<string>(dirs);
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

