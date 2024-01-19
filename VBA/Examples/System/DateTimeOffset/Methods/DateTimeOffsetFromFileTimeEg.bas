Attribute VB_Name = "DateTimeOffsetFromFileTimeEg"
'@IgnoreModule EmptyMethod
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 21, 2023
'@LastModified January 10, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.fromfiletime?view=netframework-4.8.1#examples

Option Explicit

''
' The following example retrieves the Windows file times for the WordPad executable.
''
Public Sub DateTimeOffsetFromFileTime()
    ' Open file %windir%\write.exe
    Dim winDir As DotNetLib.String
    Set winDir = Strings.Create(Environment.SystemDirectory)
    If (Not (winDir.EndsWith3(Path.DirectorySeparatorChar))) Then
        Set winDir = Strings.Concat2(winDir, Path.DirectorySeparatorChar)
    End If
    Set winDir = Strings.Concat2(winDir, "write.exe")
    
    '/ Get date and time, convert to file time, then convert back
    Dim pvtfileInfo As DotNetLib.FileInfo
    Set pvtfileInfo = FileInfo.Create(winDir.ToString())
    Dim infoCreationTime As DotNetLib.DateTimeOffset
    Dim infoAccessTime As DotNetLib.DateTimeOffset
    Dim infoWriteTime As DotNetLib.DateTimeOffset
    Dim ftCreationTime As LongLong
    Dim ftAccessTime As LongLong
    Dim ftWriteTime As LongLong
    
    ' Get dates and times of file creation, last access, and last write
    Set infoCreationTime = DateTimeOffset.CreateFromDateTime(pvtfileInfo.creationTime)
    Set infoAccessTime = DateTimeOffset.CreateFromDateTime(pvtfileInfo.lastAccessTime)
    Set infoWriteTime = DateTimeOffset.CreateFromDateTime(pvtfileInfo.lastWriteTime)
    ' Convert values to file times
    ftCreationTime = infoCreationTime.ToFileTime()
    ftAccessTime = infoAccessTime.ToFileTime()
    ftWriteTime = infoWriteTime.ToFileTime()
    
    ' Convert file times back to DateTimeOffset values
    Debug.Print VBString.Format("File {0} Retrieved Using a FileInfo Object:", winDir)
    Debug.Print VBString.Format("   Created:     {0:d}", DateTimeOffset.FromFileTime(ftCreationTime).ToString())
    Debug.Print VBString.Format("   Last Access: {0:d}", DateTimeOffset.FromFileTime(ftAccessTime).ToString())
    Debug.Print VBString.Format("   Last Write:  {0:d}", DateTimeOffset.FromFileTime(ftWriteTime).ToString())
End Sub

' The example produces the following output:
'
'    File C:\WINDOWS\system32\write.exe Retrieved Using a FileInfo Object:
'       Created:     10/13/2005 5:26:59 PM -07:00
'       Last Access: 3/20/2007 2:07:00 AM -07:00
'       Last Write:  8/4/2004 5:00:00 AM -07:00
