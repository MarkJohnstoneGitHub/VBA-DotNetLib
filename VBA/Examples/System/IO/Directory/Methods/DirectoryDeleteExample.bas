Attribute VB_Name = "DirectoryDeleteExample"
'@Folder("Examples.System.IO.Directory.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 7, 2023
'@LastModified November 7, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.delete?view=netframework-4.8.1#system-io-directory-delete(system-string)

Option Explicit

''
' The following example shows how to create a new directory and subdirectory,
' and then delete only the subdirectory.
''
Public Sub DirectoryDeleteExample()
    Dim subPath As String
    subPath = "C:\NewDirectory\NewSubDirectory"
    On Error GoTo ErrorHandler
    Call Directory.CreateDirectory(subPath)
    Call Directory.Delete(subPath)
    Dim directoryExists As Boolean
    directoryExists = Directory.Exists("C:\NewDirectory")
    Dim subDirectoryExists As Boolean
    subDirectoryExists = Directory.Exists(subPath)
    Debug.Print "top-level directory exists: " & directoryExists
    Debug.Print "sub-directory exists: " & subDirectoryExists
Exit Sub
ErrorHandler:
    Debug.Print VBAString.Format("The process failed: {0}", Err.Description)
End Sub

'@TODO https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.delete?view=netframework-4.8.1#system-io-directory-delete(system-string)
Public Sub DirectoryDeleteExample2()
    Dim topPath As String
    topPath = "C:\NewDirectory"
    Dim subPath As String
    subPath = "C:\NewDirectory\NewSubDirectory"
    On Error GoTo ErrorHandler
    Call Directory.CreateDirectory(subPath)
    
    
        '                using (StreamWriter writer = File.CreateText(subPath + @"\example.txt"))
'                {
'                    writer.WriteLine("content added");
'                }
        
Exit Sub
ErrorHandler:

End Sub

'using System;
'using System.IO;
'
'Namespace ConsoleApplication
'{
'    Class Program
'    {
'        static void Main(string[] args)
'        {
'            string topPath = @"C:\NewDirectory";
'            string subPath = @"C:\NewDirectory\NewSubDirectory";
'
'            Try
'            {
'                Directory.CreateDirectory(subPath);
'
'                using (StreamWriter writer = File.CreateText(subPath + @"\example.txt"))
'                {
'                    writer.WriteLine("content added");
'                }
'
'                Directory.Delete(topPath, true);
'
'                bool directoryExists = Directory.Exists(topPath);
'
'                Console.WriteLine("top-level directory exists: " + directoryExists);
'            }
'            catch (Exception e)
'            {
'                Console.WriteLine("The process failed: {0}", e.Message);
'            }
'        }
'    }
'}
