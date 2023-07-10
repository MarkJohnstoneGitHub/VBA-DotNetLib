Attribute VB_Name = "DateTimeFromBinaryExample"
'@Folder("VBADotNetLib.Examples.DateTime.Methods")

'@Author Mark Johnstone https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 11, 2023
'@LastModified July 11, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.frombinary?view=netframework-4.8.1

Option Explicit

Public Sub DateTimeFromBinary()
   Dim localDate As DateTime
   Set localDate = DateTime.CreateFromDateTimeKind(2023, 10, 1, 2, 30, 0, DateTimeKind.DateTimeKind_Local)
   
   Dim binLocal As LongLong
   binLocal = localDate.ToBinary()
   
   '@TODO Implement wrapper for DotNetLib.TimeZoneInfo
   Dim tzi As DotNetLib.TimeZoneInfo
   Set tzi = New DotNetLib.TimeZoneInfo
   If (tzi.Local.IsInvalidTime(localDate.ComObject)) Then
      Debug.Print localDate.ToString & " is an invalid time in the " & tzi.Local.StandardName
   End If
   Dim localDate2 As DateTime
   Set localDate2 = DateTime.FromBinary(binLocal)
   Debug.Print localDate.ToString & " = " & localDate2.ToString & ": " & localDate.Equals(localDate2)

' The example displays the following output for the local time zone AUS Eastern Standard Time:
'    1/10/2023 2:30:00 AM is an invalid time in the AUS Eastern Standard Time
'    1/10/2023 2:30:00 AM = 1/10/2023 3:30:00 AM: False
End Sub

