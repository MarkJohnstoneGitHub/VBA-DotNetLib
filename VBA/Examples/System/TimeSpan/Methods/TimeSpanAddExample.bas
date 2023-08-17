Attribute VB_Name = "TimeSpanAddExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.add?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example calls the Add method to add each element in an array of time intervals to a base TimeSpan value.")
Public Sub TimeSpanAdd()
Attribute TimeSpanAdd.VB_Description = "The following example calls the Add method to add each element in an array of time intervals to a base TimeSpan value."
   Dim intervals() As ITimeSpan
   Dim baseTimeSpan As ITimeSpan
   Set baseTimeSpan = TimeSpan.Create2(1, 12, 15, 16)
   
   ' Create an array of timespan intervals.
   Objects.ToArray intervals, _
                  TimeSpan.FromDays(1.5), _
                  TimeSpan.FromHours(1.5), _
                  TimeSpan.FromMinutes(45), _
                  TimeSpan.FromMilliseconds(505), _
                  TimeSpan.Create2(1, 17, 32, 20), _
                  TimeSpan.Create(-8, 30, 0)
                  
   ' Calculate a new time interval by adding each element to the base interval.
   Dim varInterval As Variant
   For Each varInterval In intervals
      Dim interval As ITimeSpan
      Set interval = varInterval
      Debug.Print baseTimeSpan.ToString2("g") _
            & " " & IIf(TimeSpan.LessThan(interval, TimeSpan.Zero), "-", "+") _
            & " " & varInterval.ToString2("%d\:hh\:mm\:ss\.ffff") _
            & " = " _
            & baseTimeSpan.Add(varInterval).ToString2("%d\:hh\:mm\:ss\.ffff")
   Next
End Sub

' The example displays the following output:
'       1:12:15:16 + 1:12:00:00.0000 = 3:00:15:16.0000
'       1:12:15:16 + 0:01:30:00.0000 = 1:13:45:16.0000
'       1:12:15:16 + 0:00:45:00.0000 = 1:13:00:16.0000
'       1:12:15:16 + 0:00:00:00.5050 = 1:12:15:16.5050
'       1:12:15:16 + 1:17:32:20.0000 = 3:05:47:36.0000
'       1:12:15:16 - 0:07:30:00.0000 = 1:04:45:16.0000
