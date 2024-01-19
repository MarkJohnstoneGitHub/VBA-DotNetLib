Attribute VB_Name = "PathGetFullPathExample"
'@Folder "Examples.System.IO.Path.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 20, 2023
'@LastModified November 20, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.getfullpath?view=netframework-4.8.1#system-io-path-getfullpath(system-string)

Option Explicit

''
' The following example demonstrates the GetFullPath method on a Windows-based
' desktop platform.
''
Public Sub PathGetFullPathExample()
    Dim fileName As String
    fileName = "myfile.ext"
    Dim path1 As String
    path1 = "mydir"
    Dim path2 As String
    path2 = "\mydir"
    Dim fullPath As String
    
    fullPath = Path.GetFullPath(path1)
    Debug.Print VBString.Format("GetFullPath('{0}') returns '{1}'", _
                                path1, fullPath)
                                
    fullPath = Path.GetFullPath(fileName)
    Debug.Print VBString.Format("GetFullPath('{0}') returns '{1}'", _
                                fileName, fullPath)

    fullPath = Path.GetFullPath(path2)
    Debug.Print VBString.Format("GetFullPath('{0}') returns '{1}'", _
                                path2, fullPath)
End Sub

' Output is based on your current directory, except
' in the last case, where it is based on the root drive
' GetFullPath('mydir') returns 'C:\temp\Demo\mydir'
' GetFullPath('myfile.ext') returns 'C:\temp\Demo\myfile.ext'
' GetFullPath('\mydir') returns 'C:\mydir'


