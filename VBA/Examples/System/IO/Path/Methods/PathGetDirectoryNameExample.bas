Attribute VB_Name = "PathGetDirectoryNameExample"
'@Folder("Examples.System.IO.Path.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 19, 2023
'@LastModified November 19, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.getdirectoryname?view=netframework-4.8.1#system-io-path-getdirectoryname(system-string)

Option Explicit

''
' The following example demonstrates using the GetDirectoryName method on a
' Windows-based desktop platform.
''
Public Sub PathGetDirectoryNameExample()
    Dim filePath As String
    filePath = "C:\MyDir\MySubDir\myfile.ext"
    Dim directoryName As String
    Dim i As Long
    
    Do While (filePath <> VBA.vbNullString)
        directoryName = Path.GetDirectoryName(filePath)
        Debug.Print VBAString.Format("GetDirectoryName('{0}') returns '{1}'", _
                                    filePath, directoryName)
        filePath = directoryName
        If (i = 1) Then
             filePath = directoryName + "\"  ' this will preserve the previous path
        End If
        i = i + 1
    Loop
End Sub

'/*
'This code produces the following output:
'
'GetDirectoryName('C:\MyDir\MySubDir\myfile.ext') returns 'C:\MyDir\MySubDir'
'GetDirectoryName('C:\MyDir\MySubDir') returns 'C:\MyDir'
'GetDirectoryName('C:\MyDir\') returns 'C:\MyDir'
'GetDirectoryName('C:\MyDir') returns 'C:\'
'GetDirectoryName('C:\') returns ''
'*/
