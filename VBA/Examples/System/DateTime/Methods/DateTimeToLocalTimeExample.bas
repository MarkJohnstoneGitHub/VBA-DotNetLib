Attribute VB_Name = "DateTimeToLocalTimeExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 13, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.tolocaltime?view=netframework-4.8.1#examples

Option Explicit

Private Const ParseDateTime As String = "12/10/2015 6:18 AM"
Private Const ParseUniversalDateTime As String = "12/10/2015 6:18:00"

'@Description("The following example demonstrates the ToLocalTime method.")
' Note that the exact output depends on the current culture and the local time zone of the system on which it is run.
Public Sub DateTimeToLocalTime()
Attribute DateTimeToLocalTime.VB_Description = "The following example demonstrates the ToLocalTime method."
   Dim LocalDateTime As IDateTime
   Dim univDateTime As IDateTime
   
   Dim strDateTime As String
   strDateTime = ParseDateTime
   
   On Error Resume Next    'Skip error and continue running
   Set LocalDateTime = DateTime.Parse(strDateTime)
   If Err.number = 0 Then  ' Try
      Set univDateTime = LocalDateTime.ToUniversalTime()
      Debug.Print LocalDateTime.ToString & " local time is " & univDateTime.ToString & " universal time."
   Else ' Catch
      If Err.number = COMHResult.FormatException Then
         Debug.Print "Invalid format."
      Else
         'Unexpected Error should not occur
         Debug.Print "0x" & Hex$(Err.number) & " " & Err.Description
      End If
   End If
   On Error GoTo 0 'Stop code and display error. End of Try, Catch block

   On Error Resume Next    'Skip error and continue running
   Set univDateTime = DateTime.Parse(ParseUniversalDateTime)
   If Err.number = 0 Then     ' Try
      Set LocalDateTime = univDateTime.ToLocalTime()
      Debug.Print univDateTime.ToString & " universal time is " & LocalDateTime.ToString & " local time."
   Else
      ' Catch
      If Err.number = COMHResult.FormatException Then
         Debug.Print "Invalid format."
      Else
         'Unexpected Error should not occur
         Debug.Print "0x" & Hex$(Err.number) & " " & Err.Description
      End If
   End If
   On Error GoTo 0 'Stop code and display error. End of Try, Catch block
End Sub

' Output for US Pacific time zone
'    12/10/2015 6:18:00 AM local time is 12/10/2015 2:18:00 PM universal time.
'    12/10/2015 6:18:00 AM universal time is 12/9/2015 10:18:00 PM local time.

' Output for Aus Eastern time zone
'    12/10/2015 6:18:00 AM local time is 11/10/2015 7:18:00 PM universal time.
'    12/10/2015 6:18:00 AM universal time is 12/10/2015 5:18:00 PM local time.
