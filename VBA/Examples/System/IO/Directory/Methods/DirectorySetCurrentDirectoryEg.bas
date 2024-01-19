Attribute VB_Name = "DirectorySetCurrentDirectoryEg"
'@Folder "Examples.System.IO.Directory.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 19, 2023
'@LastModified November 19, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.setcurrentdirectory?view=netframework-4.8.1#examples

Option Explicit

''
' The following example illustrates how to set the current directory and
' display the directory root.
''
Public Sub DirectorySetCurrentDirectory()
    ' Create string for a directory. This value should be an existing directory
    ' or the sample will throw a DirectoryNotFoundException.
    Dim dir As String
    dir = "C:\test"
    
    On Error GoTo ErrorHandler
    ' Set the current directory.
    Call Directory.SetCurrentDirectory(dir)
    ' Print to console the results.
    Debug.Print VBString.Format("Root directory: {0}", Directory.GetDirectoryRoot(dir))
    Debug.Print VBString.Format("Current directory: {0}", Directory.GetCurrentDirectory())
    Exit Sub
ErrorHandler:
    If Err.Number = DirectoryNotFoundException Then
        Debug.Print VBString.Format("The specified directory does not exist. {0}", Err.Description);
    Else
        Debug.Print Err.Description
    End If
End Sub

' The output of this sample depends on what value you assign to the variable dir.
' If the directory c:\test exists, the output for this sample is:
' Root directory: C:\
' Current directory: C:\test
