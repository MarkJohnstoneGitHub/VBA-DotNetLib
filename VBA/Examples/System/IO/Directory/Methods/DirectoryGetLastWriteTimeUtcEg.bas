Attribute VB_Name = "DirectoryGetLastWriteTimeUtcEg"
'@Folder "Examples.System.IO.Directory.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 19, 2023
'@LastModified November 19, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.getlastwritetimeutc?view=netframework-4.8.1#examples

Option Explicit

''
' The following example illustrates the differences in output when using
' Coordinated Universal Time (UTC) output.
''
Public Sub DirectoryGetLastWriteTimeUtc()
    ' Set the directory.
    Dim n As String
    n = "C:\test\newdir"
    ' Create two variables to use to set the time.
    Dim dtime1 As DotNetLib.DateTime
    Set dtime1 = DateTime.CreateFromDate(2002, 1, 3)
    Dim dtime2 As DotNetLib.DateTime
    Set dtime2 = DateTime.CreateFromDate(1999, 1, 1)
    
    On Error GoTo ErrorHandler
    Call Directory.CreateDirectory(n)
    
    ' Set the creation and last access times to a variable DateTime value.
    Call Directory.SetCreationTime(n, dtime1)
    Call Directory.SetLastAccessTimeUtc(n, dtime1)

    ' Print to console the results.
    Debug.Print VBString.Format("Creation Date: {0}", Directory.GetCreationTime(n))
    Debug.Print VBString.Format("UTC creation Date: {0}", Directory.GetCreationTimeUtc(n))
    Debug.Print VBString.Format("Last write time: {0}", Directory.GetLastWriteTime(n))
    Debug.Print VBString.Format("UTC last write time: {0}", Directory.GetLastWriteTimeUtc(n))
    Debug.Print VBString.Format("Last access time: {0}", Directory.GetLastAccessTime(n))
    Debug.Print VBString.Format("UTC last access time: {0}", Directory.GetLastAccessTimeUtc(n))

    ' Set the last write time to a different value.
    Call Directory.SetLastWriteTimeUtc(n, dtime2)
    Debug.Print VBString.Format("Changed last write time: {0}", Directory.GetLastWriteTimeUtc(n))

Exit Sub
ErrorHandler:
    Debug.Print VBString.Format("The process failed: {0}", Err.Description)
End Sub

' Obviously, since this sample deals with dates and times, the output will vary
' depending on when you run the executable. Here is one example of the output:
'Creation Date: 1/3/2002 12:00:00 AM
'UTC creation Date: 1/3/2002 8:00:00 AM
'Last write time: 12/31/1998 4:00:00 PM
'UTC last write time: 1/1/1999 12:00:00 AM
'Last access time: 1/2/2002 4:00:00 PM
'UTC last access time: 1/3/2002 12:00:00 AM
'Changed last write time: 1/1/1999 12:00:00 AM
