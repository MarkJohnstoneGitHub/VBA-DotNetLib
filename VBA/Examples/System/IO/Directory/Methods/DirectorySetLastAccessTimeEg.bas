Attribute VB_Name = "DirectorySetLastAccessTimeEg"
'@Folder("Examples.System.IO.Directory.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 19, 2023
'@LastModified November 19, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.setlastaccesstime?view=netframework-4.8.1#examples

Option Explicit

''
' The following example demonstrates how to use SetLastAccessTime.
''
Public Sub DirectorySetLastAccessTime()
    On Error GoTo ErrorHandler
    Dim pvtPath As String
    pvtPath = "c:\MyDir"
    If (Not Directory.Exists(pvtPath)) Then
        Call Directory.CreateDirectory(pvtPath)
    End If
    Call Directory.SetLastAccessTime(pvtPath, DateTime.CreateFromDate(1985, 5, 4))

    ' Get the last access time of a well-known directory.
    Dim dt As DotNetLib.DateTime
    Set dt = Directory.GetLastAccessTime(pvtPath)
    Debug.Print VBAString.Format("The last access time for this directory was {0}", dt)
Exit Sub
ErrorHandler:
    Debug.Print VBAString.Format("The process failed: {0}", Err.Description)
End Sub

'Example Output:
'    The last access time for this directory was 4/05/1985 12:00:00 AM
