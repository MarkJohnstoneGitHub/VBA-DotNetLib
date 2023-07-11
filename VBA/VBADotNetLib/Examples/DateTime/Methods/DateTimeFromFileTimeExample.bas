Attribute VB_Name = "DateTimeFromFileTimeExample"
'@Folder("VBADotNetLib.Examples.DateTime.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 11, 2023
'@LastModified July 11, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.fromfiletime?view=netframework-4.8.1

Option Explicit

Public Sub DateTimeFromFileTime()
   Dim date1 As DateTime
   Set date1 = DateTime.CreateFromDateTime(2023, 10, 1, 2, 30, 0)
   Debug.Print "Invalid Time: " & TimeZoneInfo.Locale.IsInvalidTime(date1)

   Dim ft As LongLong
   ft = date1.ToFileTime()
   
   Dim date2 As DateTime
   Set date2 = DateTime.FromFileTime(ft)
   Debug.Print date1.ToString & " -> " & date2.ToString
   
' The example displays the following output for local time zone of AUS Eastern Standard Time:
'       Invalid Time: True
'       1/10/2023 2:30:00 AM -> 1/10/2023 3:30:00 AM
End Sub

