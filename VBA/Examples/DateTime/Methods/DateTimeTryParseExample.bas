Attribute VB_Name = "DateTimeTryParseExample"
'@Folder("VBADotNetLib.Examples.DateTime.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 14, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.tryparse?view=netframework-4.8.1

Option Explicit

'@Description("The following example passes a number of date and time strings to the DateTime.TryParse(String, DateTime) method.")
Public Sub DateTimeTryParse()
Attribute DateTimeTryParse.VB_Description = "The following example passes a number of date and time strings to the DateTime.TryParse(String, DateTime) method."
   Dim dateStrings() As String
   dateStrings = Strings.ToArray("05/01/2009 14:57:32.8", "2009-05-01 14:57:32.8", _
                                 "2009-05-01T14:57:32.8375298-04:00", "5/01/2008", _
                                 "5/01/2008 14:57:32.80 -07:00", _
                                 "1 May 2008 2:57:32.8 PM", "16-05-2009 1:00:32 PM", _
                                 "Fri, 15 May 2009 20:10:57 GMT")
   
   'Debug.Print "Attempting to parse strings using " & CultureInfo.CurrentCulture.Name
   Dim dateValue As IDateTime
   Dim dateString As Variant
   For Each dateString In dateStrings
      If (DateTime.TryParse(dateString, dateValue)) Then
         Debug.Print "  Converted '" & dateString & "' to " & dateValue.ToString() & " (" & dateValue.Kind & ")"
      Else
         Debug.Print "  Unable to parse '" & dateString & "'."
      End If
   Next
End Sub

' The example displays output like the following:
'    Attempting to parse strings using en-US culture.
'      Converted '05/01/2009 14:57:32.8' to 5/1/2009 2:57:32 PM (Unspecified).
'      Converted '2009-05-01 14:57:32.8' to 5/1/2009 2:57:32 PM (Unspecified).
'      Converted '2009-05-01T14:57:32.8375298-04:00' to 5/1/2009 11:57:32 AM (Local).
'
'      Converted '5/01/2008' to 5/1/2008 12:00:00 AM (Unspecified).
'      Converted '5/01/2008 14:57:32.80 -07:00' to 5/1/2008 2:57:32 PM (Local).
'      Converted '1 May 2008 2:57:32.8 PM' to 5/1/2008 2:57:32 PM (Unspecified).
'      Unable to parse '16-05-2009 1:00:32 PM'.
'      Converted 'Fri, 15 May 2009 20:10:57 GMT' to 5/15/2009 1:10:57 PM (Local).
