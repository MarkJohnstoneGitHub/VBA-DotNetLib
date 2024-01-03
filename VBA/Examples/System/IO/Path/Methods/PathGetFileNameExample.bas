Attribute VB_Name = "PathGetFileNameExample"
'@Folder("Examples.System.IO.Path.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 20, 2023
'@LastModified November 20, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.getfilename?view=netframework-4.8.1#system-io-path-getfilename(system-string)

Option Explicit

''
' The following example demonstrates the behavior of the GetFileName method
' on a Windows-based desktop platform.
''
Public Sub PathGetFileNameExample()
    Dim filename As String
    filename = "C:\mydir\myfile.ext"
    Dim pvtPath As String
    pvtPath = "C:\mydir\"
    Dim result As String
    
    result = Path.GetFileName(filename)
    Debug.Print VBAString.Format("GetFileName('{0}') returns '{1}'", _
                                filename, result)
                                
    result = Path.GetFileName(pvtPath)
    Debug.Print VBAString.Format("GetFileName('{0}') returns '{1}'", _
                                pvtPath, result)
End Sub
' This code produces output similar to the following:
'
' GetFileName('C:\mydir\myfile.ext') returns 'myfile.ext'
' GetFileName('C:\mydir\') returns ''
