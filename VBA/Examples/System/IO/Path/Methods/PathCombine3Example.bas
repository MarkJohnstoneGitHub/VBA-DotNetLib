Attribute VB_Name = "PathCombine3Example"
'@Folder "Examples.System.IO.Path.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 19, 2023
'@LastModified November 19, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.combine?view=netframework-4.8.1#system-io-path-combine(system-string-system-string-system-string)

Option Explicit

''
' The following example combines three paths.
''
Public Sub PathCombine3Example()
    Dim p1 As String
    p1 = "d:\archives\"
    Dim p2 As String
    p2 = "media"
    Dim p3 As String
    p3 = "images"
    Dim combined As String
    combined = Path.Combine3(p1, p2, p3)
    Debug.Print combined
End Sub

'Example Output:
'   d:\archives\media\images
