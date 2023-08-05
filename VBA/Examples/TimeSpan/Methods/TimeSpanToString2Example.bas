Attribute VB_Name = "TimeSpanToString2Example"
'@Folder("VBADotNetLib.Examples.TimeSpan.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 4, 2023
'@LastModified August 4, 2023

'@Reference

Option Explicit

'@Description("The following example uses standard and custom TimeSpan format strings to display the string representation of each element in an array of TimeSpan values.")
'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.tostring?view=netframework-4.8.1#system-timespan-tostring(system-string)
Public Sub TimeSpanToString2()
Attribute TimeSpanToString2.VB_Description = "The following example uses standard and custom TimeSpan format strings to display the string representation of each element in an array of TimeSpan values."
   ' Create an array of timespan intervals.
   Dim spans() As ITimeSpan
   Objects.ToArray spans, _
      TimeSpan.Zero, _
      TimeSpan.Create3(-14, 0, 0, 0, 0), _
      TimeSpan.Create(1, 2, 3), _
      TimeSpan.Create3(0, 0, 0, 0, 250), _
      TimeSpan.Create3(99, 23, 59, 59, 999), _
      TimeSpan.Create(3, 0, 0), _
      TimeSpan.Create3(0, 0, 0, 0, 25)

   Dim fmts() As String
   fmts = Strings.ToArray("c", "g", "G", "hh\:mm\:ss", "%m' min.'")
   
   ' Calculate a new time interval by adding each element to the base interval.
   Dim varSpan As Variant
   For Each varSpan In spans
      Dim span As ITimeSpan
      Set span = varSpan
      
      Dim fmt As Variant
      For Each fmt In fmts
         Debug.Print fmt & ": " & span.ToString2(fmt)
      Next
      Debug.Print
   Next
End Sub

' The example displays the following output:
'       c: 00:00:00
'       g: 0:00:00
'       G: 0:00:00:00.0000000
'       hh\:mm\:ss: 00:00:00
'       %m' min.': 0 min.
'
'       c: -14.00:00:00
'       g: -14:0:00:00
'       G: -14:00:00:00.0000000
'       hh\:mm\:ss: 00:00:00
'       %m' min.': 0 min.
'
'       c: 01:02:03
'       g: 1:02:03
'       G: 0:01:02:03.0000000
'       hh\:mm\:ss: 01:02:03
'       %m' min.': 2 min.
'
'       c: 00:00:00.2500000
'       g: 0:00:00.25
'       G: 0:00:00:00.2500000
'       hh\:mm\:ss: 00:00:00
'       %m' min.': 0 min.
'
'       c: 99.23:59:59.9990000
'       g: 99:23:59:59.999
'       G: 99:23:59:59.9990000
'       hh\:mm\:ss: 23:59:59
'       %m' min.': 59 min.
'
'       c: 03:00:00
'       g: 3:00:00
'       G: 0:03:00:00.0000000
'       hh\:mm\:ss: 03:00:00
'       %m' min.': 0 min.
'
'       c: 00:00:00.0250000
'       g: 0:00:00.025
'       G: 0:00:00:00.0250000
'       hh\:mm\:ss: 00:00:00
'       %m' min.': 0 min.
