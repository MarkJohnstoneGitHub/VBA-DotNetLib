Attribute VB_Name = "DirectoryGetCreationTimeUtcEg"
'@Folder("Examples.System.IO.Directory.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 17 2023
'@LastModified November 17, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.getcreationtimeutc?view=netframework-4.8.1#examples

Option Explicit

Public Sub DirectoryGetCreationTimeUtc()
    ' Set the directory.
    Dim n As String
    n = "C:\test\newdir"
    ' Create two variables to use to set the time.
    Dim dtime1 As DotNetLib.DateTime
    Set dtime1 = DateTime.CreateFromDate(2002, 1, 3)
    Dim dtime2 As DotNetLib.DateTime
    Set dtime2 = DateTime.CreateFromDate(1999, 1, 1)

    On Error Resume Next
    Directory.CreateDirectory (n)
    If Err.number = IOException Then
        Debug.Print Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
    
    ' Set the creation and last access times to a variable DateTime value.
    Call Directory.SetCreationTime(n, dtime1)
    Call Directory.SetLastAccessTimeUtc(n, dtime1)

    ' Print to console the results.
    Debug.Print VBAString.Format("Creation Date: {0}", Directory.GetCreationTime(n))
    Debug.Print VBAString.Format("UTC creation Date: {0}", Directory.GetCreationTimeUtc(n))
    Debug.Print VBAString.Format("Last write time: {0}", Directory.GetLastWriteTime(n))
    Debug.Print VBAString.Format("UTC last write time: {0}", Directory.GetLastWriteTimeUtc(n))
    Debug.Print VBAString.Format("Last access time: {0}", Directory.GetLastAccessTime(n))
    Debug.Print VBAString.Format("UTC last access time: {0}", Directory.GetLastAccessTimeUtc(n))

    ' Set the last write time to a different value.
    Call Directory.SetLastWriteTimeUtc(n, dtime2)
    Debug.Print VBAString.Format("Changed last write time: {0}", Directory.GetLastWriteTimeUtc(n))

End Sub

' Obviously, since this sample deals with dates and times, the output will vary
' depending on when you run the executable. Here is one example of the output:
' Creation Date: 3/01/2002 12:00:00 AM
' UTC creation Date: 2/01/2002 1:00:00 PM
' Last write time: 17/11/2023 8:06:08 PM
' UTC last write time: 17/11/2023 9:06:08 AM
' Last access time: 3/01/2002 11:00:00 AM
' UTC last access time: 3/01/2002 12:00:00 AM
' Changed last write time: 1/01/1999 12:00:00 AM
