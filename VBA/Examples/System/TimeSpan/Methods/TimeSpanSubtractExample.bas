Attribute VB_Name = "TimeSpanSubtractExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 17, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.subtract?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses the Subtract method to calculate the difference
' between a single TimeSpan value and each of the time intervals in an array.
' Note that, because TimeSpan format strings do not include negative signs in
' the result string, the example uses conditional logic to include a negative
' sign with negative time intervals.
''
Public Sub TimeSpanSubtract()
    Dim baseTimeSpan As DotNetLib.TimeSpan
    Set baseTimeSpan = TimeSpan.Create2(1, 12, 15, 16)
    
    ' Create an array of timespan intervals.
    Dim intervals() As DotNetLib.TimeSpan
    ObjectArray.CreateInitialize1D intervals, _
        TimeSpan.FromDays(1.5), _
        TimeSpan.FromHours(1.5), _
        TimeSpan.FromMinutes(45), _
        TimeSpan.FromMilliseconds(505), _
        TimeSpan.Create2(1, 17, 32, 20), _
        TimeSpan.Create(-8, 30, 0)
 
    ' Calculate a new time interval by adding each element to the base interval.
    Dim varInterval As Variant
    For Each varInterval In intervals
        Dim interval As DotNetLib.TimeSpan
        Set interval = varInterval
      
        Debug.Print VBString.Format("{0,-10:g} - {3}{1,15:%d\:hh\:mm\:ss\.ffff} = {4}{2:%d\:hh\:mm\:ss\.ffff}", _
                baseTimeSpan, interval, baseTimeSpan.Subtract(interval), _
                IIf(TimeSpan.LessThan(interval, TimeSpan.Zero), "-", VBA.vbNullString), _
                IIf(TimeSpan.LessThan(baseTimeSpan, interval.Duration()), "-", VBA.vbNullString))
   Next
End Sub

' The example displays the following output:
'       1:12:15:16 - 1:12:00:00.0000 = 0:00:15:16.0000
'       1:12:15:16 - 0:01:30:00.0000 = 1:10:45:16.0000
'       1:12:15:16 - 0:00:45:00.0000 = 1:11:30:16.0000
'       1:12:15:16 - 0:00:00:00.5050 = 1:12:15:15.4950
'       1:12:15:16 - 1:17:32:20.0000 = -0:05:17:04.0000
'       1:12:15:16 - -0:07:30:00.0000 = 1:19:45:16.0000

