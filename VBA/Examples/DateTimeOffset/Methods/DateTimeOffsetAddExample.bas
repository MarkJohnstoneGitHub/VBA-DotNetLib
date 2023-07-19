Attribute VB_Name = "DateTimeOffsetAddExample"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 20, 2023
'@LastModified July 20, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.add?view=netframework-4.8.1#examples

Option Explicit

' The following example creates an array of TimeSpan objects that represent the flight times between destinations.
' The Add method then adds these times to a DateTimeOffset object that represents a flight's initial takeoff time.
' The result reflects the scheduled arrival time at each destination.
Public Sub DateTimeOffsetAdd()
   Dim takeOff As DateTimeOffset
   Set takeOff = DateTimeOffset.CreateFromDateTimeParts(2007, 6, 1, 7, 55, 0, TimeSpan.Create(-5, 0, 0))
   
   Dim currentTime As DateTimeOffset
   Set currentTime = takeOff
   
   Dim flightTimes() As TimeSpan
   Objects.ToArray flightTimes, TimeSpan.Create(2, 25, 0), TimeSpan.Create(1, 48, 0)
   
   Debug.Print "Takeoff is scheduled for " & takeOff.ToString2("d") & " at " & takeOff.ToString2("T") & "."
   
   Dim ctr As Long
   
   For ctr = LBound(flightTimes) To UBound(flightTimes)
      Set currentTime = currentTime.Add(flightTimes(ctr))
      Debug.Print "Destination #" & ctr + 1 & " at " & currentTime.ToString()
   Next

' Output:
'   Takeoff is scheduled for 1/06/2007 at 7:55:00 AM.
'   Destination #1 at 1/06/2007 10:20:00 AM -05:00
'   Destination #2 at 1/06/2007 12:08:00 PM -05:00
End Sub
