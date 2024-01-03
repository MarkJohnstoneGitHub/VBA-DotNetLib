Attribute VB_Name = "PathGetExtensionExample"
'@Folder("Examples.System.IO.Path.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 20, 2023
'@LastModified November 20, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.getextension?view=netframework-4.8.1#system-io-path-getextension(system-string)

Option Explicit

''
' The following example demonstrates using the GetExtension method on a
' Windows-based desktop platform.
''
Public Sub PathGetExtensionExample()
    Dim filename As String
    filename = "C:\mydir.old\myfile.ext"
    Dim pvtPath As String
    pvtPath = "C:\mydir.old\"
    Dim extension As String
    extension = Path.GetExtension(filename)
    Debug.Print VBAString.Format("GetExtension('{0}') returns '{1}'", _
                                pvtPath, extension)
    extension = Path.GetExtension(pvtPath)
    Debug.Print VBAString.Format("GetExtension('{0}') returns '{1}'", _
                                pvtPath, extension)
End Sub

' This code produces output similar to the following:
'
' GetExtension('C:\mydir.old\myfile.ext') returns '.ext'
' GetExtension('C:\mydir.old\') returns ''
