Attribute VB_Name = "PathCombine4Example"
'@Folder "Examples.System.IO.Path.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 19, 2023
'@LastModified November 19, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.combine?view=netframework-4.8.1#system-io-path-combine(system-string-system-string-system-string-system-string)

Option Explicit

''
' The following example combines four paths.
''
Public Sub PathCombine4Example()
    Dim path1 As String
    path1 = "d:\archives\"
    Dim path2 As String
    path2 = "2001"
    Dim path3 As String
    path3 = "media"
    Dim path4 As String
    path4 = "images"
    
    Dim combinedPath As String
    combinedPath = Path.Combine4(path1, path2, path3, path4)
    Debug.Print combinedPath
End Sub

'Example Output:
'   d:\archives\2001\media\images
