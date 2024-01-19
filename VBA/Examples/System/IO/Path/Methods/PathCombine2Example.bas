Attribute VB_Name = "PathCombine2Example"
'@Folder "Examples.System.IO.Path.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 19, 2023
'@LastModified November 19, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.combine?view=netframework-4.8.1#system-io-path-combine(system-string-system-string)

Option Explicit

''
' The following example demonstrates using the Combine method on Windows.
''
Public Sub PathCombine2Example()
    Dim path1 As String
    path1 = "c:\temp"
    Dim path2 As String
    path2 = "subdir\file.txt"
    Dim path3 As String
    path3 = "c:\temp.txt"
    Dim path4 As String
    path4 = "c:^*&)(_=@#'\^&#2.*(.txt"
    Dim path5 As String
    path5 = ""
    Call CombinePaths(path1, path2)
    Call CombinePaths(path1, path3)
    Call CombinePaths(path3, path2)
    Call CombinePaths(path4, path2)
    Call CombinePaths(path5, path2)
End Sub

Private Sub CombinePaths(ByVal p1 As String, ByVal p2 As String)
    Dim combination As String
    combination = Path.Combine2(p1, p2)
    Debug.Print VBString.Format("When you combine '{0}' and '{1}', the result is: {2}'{3}'", _
                                p1, p2, Environment.NewLine, combination)
    Debug.Print
End Sub

' This code produces output similar to the following:
'
' When you combine 'c:\temp' and 'subdir\file.txt', the result is:
' 'c:\temp\subdir\file.txt'
'
' When you combine 'c:\temp' and 'c:\temp.txt', the result is:
' 'c:\temp.txt'
'
' When you combine 'c:\temp.txt' and 'subdir\file.txt', the result is:
' 'c:\temp.txt\subdir\file.txt'
'
' When you combine 'c:^*&)(_=@#'\^&#2.*(.txt' and 'subdir\file.txt', the result is:
' 'c:^*&)(_=@#'\^&#2.*(.txt\subdir\file.txt'
'
' When you combine '' and 'subdir\file.txt', the result is:
' 'subdir\file.txt'


