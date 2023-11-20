Attribute VB_Name = "PathIsPathRootedExample"
'@Folder("Examples.System.IO.Path.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 20, 2023
'@LastModified November 20, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.ispathrooted?view=netframework-4.8.1#system-io-path-ispathrooted(system-string)

Option Explicit

''
' The following example demonstrates how the IsPathRooted method can be used to
' test three strings.
''
Public Sub PathIsPathRootedExample()
    Dim fileName As String
    fileName = "C:\mydir\myfile.ext"
    Dim UncPath As String
    UncPath = "\\myPc\mydir\myfile"
    Dim relativePath As String
    relativePath = "mydir\sudir\"
    Dim result As Boolean
    
    result = Path.IsPathRooted(fileName)
    Debug.Print VBAString.Format("IsPathRooted('{0}') returns {1}", _
                                fileName, result)
                                
    result = Path.IsPathRooted(UncPath)
    Debug.Print VBAString.Format("IsPathRooted('{0}') returns {1}", _
                                UncPath, result)

    result = Path.IsPathRooted(relativePath)
    Debug.Print VBAString.Format("IsPathRooted('{0}') returns {1}", _
                                relativePath, result)
End Sub

' This code produces output similar to the following:
'
' IsPathRooted('C:\mydir\myfile.ext') returns True
' IsPathRooted('\\myPc\mydir\myfile') returns True
' IsPathRooted('mydir\sudir\') returns False


Public Sub PathIsPathRootedExample2()
    Dim relative1 As String
    relative1 = "C:Documents"
    Call ShowPathInfo(relative1)
    
    Dim relative2 As String
    relative2 = "/Documents"
    Call ShowPathInfo(relative2)
    
    Dim absolute As String
    absolute = "C:/Documents"
    Call ShowPathInfo(absolute)
End Sub

Private Sub ShowPathInfo(ByVal pPath As String)
    Debug.Print VBAString.Format("Path: {0}", pPath)
    Debug.Print VBAString.Format("   Rooted: {0}", Path.IsPathRooted(pPath))
    Debug.Print VBAString.Format("   Full path: {0}", Path.GetFullPath(pPath))
    Debug.Print
End Sub

' The example displays the following output when run on a Windows system:
'    Path: C:Documents
'        Rooted: True
'        Full path: c:\Users\user1\Documents\projects\path\ispathrooted\Documents
'
'    Path: /Documents
'       Rooted: True
'       Full path: c:\Documents
'
'    Path: C:/Documents
'       Rooted: True
'       Full path: C:\Documents
