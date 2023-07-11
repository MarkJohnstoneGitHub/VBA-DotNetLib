Attribute VB_Name = "DateTimeParseExample"
'@Folder("VBADotNetLib.Examples.DateTime.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 12, 2023
'@LastModified July 12, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.parse?view=netframework-4.8.1#system-datetime-parse(system-string)

'@Notes
' https://stackoverflow.com/questions/44638867/vba-excel-try-catch
' FormatException Err.Number -2146233033 See enum ResourceKey

Option Explicit

Public Sub DateTimeParse()
   ' Assume the current culture is en-US.
   ' The date is February 16, 2008, 12 hours, 15 minutes and 12 seconds.
   '
   ' Use standard en-US date and time value
   Dim dateValue As DateTime
   Dim dateString As String
   dateString = "2/16/2008 12:15:12 PM"
   
   On Error Resume Next
   Set dateValue = DateTime.Parse(dateString)
   If Err.Number = 0 Then
      Debug.Print "'" & dateString & "' converted to " & dateValue.ToString & "."
   Else
      If Err.Number = Arg_FormatException Then
         Debug.Print "Unable to convert '" & dateString & "'."
      Else
         Debug.Print Err.Number, Err.Description
      End If
   End If
   On Error GoTo 0 'reset error handling
   
   ' Reverse month and day to conform to the fr-FR culture.
   ' The date is February 16, 2008, 12 hours, 15 minutes and 12 seconds.
   dateString = "16/02/2008 12:15:12"
   On Error Resume Next
   Set dateValue = DateTime.Parse(dateString)
   If Err.Number = 0 Then
      Debug.Print "'" & dateString & "' converted to " & dateValue.ToString & "."
   Else
      If Err.Number = Arg_FormatException Then
         Debug.Print "Unable to convert '" & dateString & "'."
      Else
         Debug.Print Err.Number, Err.Description
      End If
   End If
   On Error GoTo 0 'reset error handling
   
' The example displays the following output to the console:
'       '2/16/2008 12:15:12 PM' converted to 2/16/2008 12:15:12 PM.
'       Unable to convert '16/02/2008 12:15:12'.
End Sub


