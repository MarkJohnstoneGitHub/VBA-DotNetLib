Attribute VB_Name = "DateTimeParseExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 12, 2023
'@LastModified September 9, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.parse?view=netframework-4.8.1#system-datetime-parse(system-string)

'@Notes
' https://stackoverflow.com/questions/44638867/vba-excel-try-catch
' FormatException Err.Number -2146233033 See enum COMHResult

Option Explicit

' The following example parses the string representation of several date and time values by:
'
' Using the default format provider, which provides the formatting conventions of the current
' culture of the computer used to produce the example output. The output from this example
' reflects the formatting conventions of the en-US culture.
'
' Using the default style value, which is AllowWhiteSpaces.
'
' It handles the FormatException exception that is thrown when the method tries to parse the
' string representation of a date and time by using some other culture's formatting conventions.
' It also shows how to successfully parse a date and time value that does not use the formatting
' conventions of the current culture.
Public Sub DateTimeParse()
   ' Assume the current culture is en-US.
   ' The date is February 16, 2008, 12 hours, 15 minutes and 12 seconds.
   '
   ' Use standard en-US date and time value
   Dim dateValue As IDateTime
   Dim dateString As String
   dateString = "2/16/2008 12:15:12 PM"
   
   On Error Resume Next
   Set dateValue = DateTime.Parse(dateString)
   If Err.number = 0 Then
      Debug.Print "'" & dateString & "' converted to " & dateValue.ToString & "."
   Else
      If Err.number = COMHResult.FormatException Then
         Debug.Print "Unable to convert '" & dateString & "'."
      Else
         Debug.Print Err.number, Err.Description
      End If
   End If
   On Error GoTo 0 'reset error handling
   
   ' Reverse month and day to conform to the fr-FR culture.
   ' The date is February 16, 2008, 12 hours, 15 minutes and 12 seconds.
   dateString = "16/02/2008 12:15:12"
   On Error Resume Next
   Set dateValue = DateTime.Parse(dateString)
   If Err.number = 0 Then
      Debug.Print "'" & dateString & "' converted to " & dateValue.ToString & "."
   Else
      If Err.number = COMHResult.FormatException Then
         Debug.Print "Unable to convert '" & dateString & "'."
      Else
         Debug.Print Err.number, Err.Description
      End If
   End If
   On Error GoTo 0 'reset error handling
End Sub

' The example displays the following output to the console:
'       '2/16/2008 12:15:12 PM' converted to 2/16/2008 12:15:12 PM.
'       Unable to convert '16/02/2008 12:15:12'.
