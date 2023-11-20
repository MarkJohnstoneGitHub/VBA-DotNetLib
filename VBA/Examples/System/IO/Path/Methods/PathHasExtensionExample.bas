Attribute VB_Name = "PathHasExtensionExample"
'@Folder("Examples.System.IO.Path.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 20, 2023
'@LastModified November 20, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.hasextension?view=netframework-4.8.1#system-io-path-hasextension(system-string)

Option Explicit

''
' The following example demonstrates the use of the HasExtension method.
''
Public Sub PathHasExtensionExample()
    Dim fileName1 As String
    fileName1 = "myfile.ext"
    Dim fileName2 As String
    fileName2 = "mydir\myfile"
    Dim pvtPath As String
    pvtPath = "C:\mydir.ext\"
    Dim result As Boolean
    
    result = Path.HasExtension(fileName1)
    Debug.Print VBAString.Format("HasExtension('{0}') returns {1}", _
                                fileName1, result)
    
    result = Path.HasExtension(fileName2)
    Debug.Print VBAString.Format("HasExtension('{0}') returns {1}", _
                                fileName2, result)
                                
    result = Path.HasExtension(pvtPath)
    Debug.Print VBAString.Format("HasExtension('{0}') returns {1}", _
                                pvtPath, result)
End Sub

' This code produces output similar to the following:
'
' HasExtension('myfile.ext') returns True
' HasExtension('mydir\myfile') returns False
' HasExtension('C:\mydir.ext\') returns False
