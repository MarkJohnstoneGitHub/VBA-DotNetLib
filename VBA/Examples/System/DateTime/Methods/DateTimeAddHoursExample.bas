Attribute VB_Name = "DateTimeAddHoursExample"
'Rubberduck annotations
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 9, 2023
'@LastModified January 6, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.addhours?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example uses the AddHours method to add a number of whole and fractional values to a date and time. It also illustrates the loss of precision caused by passing the method a value that includes a fractional component.")
Public Sub DateTimeAddHours()
Attribute DateTimeAddHours.VB_Description = "The following example uses the AddHours method to add a number of whole and fractional values to a date and time. It also illustrates the loss of precision caused by passing the method a value that includes a fractional component."
    Dim pvtHours() As Double
    pvtHours = DoubleArray.CreateInitialize1D(0.08333, 0.16667, 0.25, 0.33333, 0.5, 0.66667, 1, 2, 29, 30, 31, 90, 365)
    Dim dateValue As DotNetLib.DateTime
    Set dateValue = DateTime.CreateFromDateTime(2009, 3, 1, 12, 0, 0)
   
    Dim varHour As Variant
    For Each varHour In pvtHours
        Debug.Print VBString.Format("{0} + {1} hour(s) = {2}", dateValue, varHour, _
                           dateValue.AddHours(varHour))
    Next
End Sub

' The example displays the following output on a system whose current
' culture is en-US:
'    3/1/2009 12:00:00 PM + 0.08333 hour(s) = 3/1/2009 12:04:59 PM
'    3/1/2009 12:00:00 PM + 0.16667 hour(s) = 3/1/2009 12:10:00 PM
'    3/1/2009 12:00:00 PM + 0.25 hour(s) = 3/1/2009 12:15:00 PM
'    3/1/2009 12:00:00 PM + 0.33333 hour(s) = 3/1/2009 12:19:59 PM
'    3/1/2009 12:00:00 PM + 0.5 hour(s) = 3/1/2009 12:30:00 PM
'    3/1/2009 12:00:00 PM + 0.66667 hour(s) = 3/1/2009 12:40:00 PM
'    3/1/2009 12:00:00 PM + 1 hour(s) = 3/1/2009 1:00:00 PM
'    3/1/2009 12:00:00 PM + 2 hour(s) = 3/1/2009 2:00:00 PM
'    3/1/2009 12:00:00 PM + 29 hour(s) = 3/2/2009 5:00:00 PM
'    3/1/2009 12:00:00 PM + 30 hour(s) = 3/2/2009 6:00:00 PM
'    3/1/2009 12:00:00 PM + 31 hour(s) = 3/2/2009 7:00:00 PM
'    3/1/2009 12:00:00 PM + 90 hour(s) = 3/5/2009 6:00:00 AM
'    3/1/2009 12:00:00 PM + 365 hour(s) = 3/16/2009 5:00:00 PM
