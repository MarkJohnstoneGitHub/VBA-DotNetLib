Attribute VB_Name = "PathChangeExtensionExample"
'@Folder "Examples.System.IO.Path.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 19, 2023
'@LastModified November 19, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.changeextension?view=netframework-4.8.1#examples

Option Explicit

''
' The following example demonstrates a use of the ChangeExtension method.
''
Public Sub PathChangeExtensionExample()
    Dim goodFileName As String
    goodFileName = "C:\mydir\myfile.com.extension"
    Dim badFileName As String
    badFileName = "C:\mydir\"
    Dim result As String
    
    result = Path.ChangeExtension(goodFileName, ".old")
    Debug.Print VBString.Format("ChangeExtension({0}, '.old') returns '{1}'", _
                                goodFileName, result)
    
    result = Path.ChangeExtension(goodFileName, "")
    Debug.Print VBString.Format("ChangeExtension({0}, '') returns '{1}'", _
                                goodFileName, result)
    
    result = Path.ChangeExtension(badFileName, ".old")
    Debug.Print VBString.Format("ChangeExtension({0}, '.old') returns '{1}'", _
                                badFileName, result)
End Sub

' This code produces output similar to the following:
'
' ChangeExtension(C:\mydir\myfile.com.extension, '.old') returns 'C:\mydir\myfile.com.old'
' ChangeExtension(C:\mydir\myfile.com.extension, '') returns 'C:\mydir\myfile.com.'
' ChangeExtension(C:\mydir\, '.old') returns 'C:\mydir\.old'


