Attribute VB_Name = "PathPathSeparatorExample"
'@Folder("Examples.System.IO.Path.Fields")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 19, 2023
'@LastModified November 19, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.path.pathseparator?view=netframework-4.8.1#examples

Option Explicit

''
' The following example demonstrates the use of the PathSeparator field.
''
Public Sub PathPathSeparatorExample()
    Debug.Print VBAString.Format("Path.AltDirectorySeparatorChar={0}", Path.AltDirectorySeparatorChar)
    Debug.Print VBAString.Format("Path.DirectorySeparatorChar={0}", Path.DirectorySeparatorChar)
    Debug.Print VBAString.Format("Path.PathSeparator={0}", Path.PathSeparator)
    Debug.Print VBAString.Format("Path.VolumeSeparatorChar={0}", Path.VolumeSeparatorChar)
    
    Debug.Print "Path.GetInvalidPathChars()="
    Dim pvtChar As Variant
    For Each pvtChar In Path.GetInvalidPathChars()
        Debug.Print pvtChar;
    Next
    Debug.Print
End Sub

' This code produces output similar to the following:
' Note that the InvalidPathCharacters contain characters
' outside of the printable character set.
'
' Path.AltDirectorySeparatorChar=/
' Path.DirectorySeparatorChar=\
' Path.PathSeparator=;
' Path.VolumeSeparatorChar=:
