Attribute VB_Name = "DirectoryGetLastWriteTimeEg"
'@Folder("Examples.System.IO.Directory.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 19, 2023
'@LastModified November 19, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.getlastwritetime?view=netframework-4.8.1#examples

Option Explicit

''
' The following example demonstrates how to use GetLastWriteTime.
''
Public Sub DirectoryGetLastWriteTime()
    On Error GoTo ErrorHandler
    Dim pvtPath As String
    pvtPath = "c:\MyDir"
    If (Not Directory.Exists(pvtPath)) Then
        Call Directory.CreateDirectory(pvtPath)
    Else
        ' Take an action which will affect the write time.
        Call Directory.SetLastWriteTime(pvtPath, DateTime.CreateFromDate(1985, 4, 3))
    End If
    
    ' Get the creation time of a well-known directory.
    Dim dt As DotNetLib.DateTime
    Set dt = Directory.GetLastWriteTime(pvtPath)
    Debug.Print VBAString.Format("The last write time for this directory was {0}", dt)
    
    ' Update the last write time.
    Call Directory.SetLastWriteTime(pvtPath, DateTime.Now)
    Set dt = Directory.GetLastWriteTime(pvtPath)
    Debug.Print VBAString.Format("The last write time for this directory was {0}", dt)
Exit Sub
ErrorHandler:
    Debug.Print VBAString.Format("The process failed: {0}", Err.Description)
End Sub

'Example Output::
'    The last write time for this directory was 3/04/1985 12:00:00 AM
'    The last write time for this directory was 19/11/2023 12:41:51 AM
