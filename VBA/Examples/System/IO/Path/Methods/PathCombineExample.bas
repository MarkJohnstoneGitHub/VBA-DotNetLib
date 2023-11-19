Attribute VB_Name = "PathCombineExample"
'@Folder("Examples.System.IO.Path.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 19, 2023
'@LastModified November 19, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.combine?view=netframework-4.8.1#system-io-path-combine(system-string())

Option Explicit

''
' The following example combines an array of strings into a path.
''
Public Sub PathCombineExample()
    Dim pvtPaths() As String
    Call ArrayEx.CreateInitialize1D(pvtPaths, "d:\archives", "2001", "media", "images")
    Dim fullPath As String
    fullPath = Path.Combine(pvtPaths)
    Debug.Print fullPath
End Sub

'Example Output:
'   d:\archives\2001\media\images
