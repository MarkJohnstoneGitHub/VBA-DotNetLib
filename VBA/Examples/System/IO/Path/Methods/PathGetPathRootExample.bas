Attribute VB_Name = "PathGetPathRootExample"
'@Folder "Examples.System.IO.Path.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 20, 2023
'@LastModified November 20, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.getpathroot?view=netframework-4.8.1#system-io-path-getpathroot(system-string)

Option Explicit

''
' The following example demonstrates a use of the GetPathRoot method.
''
Public Sub PathGetPathRootExample()
    Dim pvtPath As String
    pvtPath = "\mydir\"
    Dim fileName As String
    fileName = "myfile.ext"
    Dim fullPath As String
    fullPath = "C:\mydir\myfile.ext"
    Dim pathRoot As String
    pathRoot = Path.GetPathRoot(pvtPath)
    
    Debug.Print VBString.Format("GetPathRoot('{0}') returns '{1}'", _
                                pvtPath, pathRoot)
    
    pathRoot = Path.GetPathRoot(fileName)
    Debug.Print VBString.Format("GetPathRoot('{0}') returns '{1}'", _
                                pvtPath, pathRoot)
                                
    pathRoot = Path.GetPathRoot(fullPath)
    Debug.Print VBString.Format("GetPathRoot('{0}') returns '{1}'", _
                                fullPath, pathRoot)
End Sub

' This code produces output similar to the following:
'
' GetPathRoot('\mydir\') returns '\'
' GetPathRoot('myfile.ext') returns ''
' GetPathRoot('C:\mydir\myfile.ext') returns 'C:\'


