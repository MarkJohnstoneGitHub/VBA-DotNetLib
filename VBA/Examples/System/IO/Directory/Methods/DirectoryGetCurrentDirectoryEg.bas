Attribute VB_Name = "DirectoryGetCurrentDirectoryEg"
'@IgnoreModule FunctionReturnValueDiscarded
'@Folder "Examples.System.IO.Directory.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 17, 2023
'@LastModified January 29, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.getcurrentdirectory?view=netframework-4.8.1#examples

Option Explicit

''
' The following example demonstrates how to use the GetCurrentDirectory method.
''
Public Sub DirectoryGetCurrentDirectory()
    On Error GoTo ErrorHandler
    Dim pvtPath As String
    pvtPath = Directory.GetCurrentDirectory()
    Dim target As String
    target = "c:\temp"
    Debug.Print VBString.Format("The current directory is {0}", pvtPath)
    If (Not Directory.Exists(target)) Then
        Call Directory.CreateDirectory(target)
    End If

    ' Change the current directory.
    Environment.CurrentDirectory = target
    If (pvtPath = Directory.GetCurrentDirectory()) Then
        Debug.Print "You are in the temp directory."
    Else
        Debug.Print "You are not in the temp directory."
    End If
Exit Sub
ErrorHandler:
    Debug.Print VBString.Format("The process failed: {0}", Err.Description)
End Sub
