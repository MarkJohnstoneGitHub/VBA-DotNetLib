Attribute VB_Name = "DateTimeTimeOfDayExample"
'Rubberduck annotations
'@Folder "Examples.System.DateTime.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.timeofday?view=netframework-4.8.1#examples

'@Notes
' https://learn.microsoft.com/en-us/dotnet/standard/base-types/standard-timespan-format-strings
' https://learn.microsoft.com/en-us/dotnet/standard/base-types/custom-timespan-format-strings

Option Explicit

''
' The following example displays the value of the TimeOfDay property for an
' array of DateTime values. It also contrasts the return value with the string
' returned by the "t" standard format string in a composite formatting operation.
''
Public Sub DateTimeTimeOfDay()
    Dim dates() As DotNetLib.DateTime
    Call VBArray.CreateInitialize1D(dates, DateTime.Now, _
                            DateTime.CreateFromDateTime(2013, 9, 14, 9, 28, 0), _
                            DateTime.CreateFromDateTime(2011, 5, 28, 10, 35, 0), _
                            DateTime.CreateFromDateTime(1979, 12, 25, 14, 30, 0))
    Dim varDateTime As Variant
    For Each varDateTime In dates
        Dim dtObject As DotNetLib.DateTime
        Set dtObject = varDateTime
        Debug.Print VBString.Format("Day: {0:d} Time: {1:g}", dtObject.Date, dtObject.TimeOfDay)
        Debug.Print VBString.Format(VBString.Unescape("Day: {0:d} Time: {0:t}\n"), dtObject)
    Next
End Sub

' The example displays output like the following:
'    Day: 7/25/2012 Time: 10:08:12.9713744
'    Day: 7/25/2012 Time: 10:08 AM
'
'    Day: 9/14/2013 Time: 9:28:00
'    Day: 9/14/2013 Time: 9:28 AM
'
'    Day: 5/28/2011 Time: 10:35:00
'    Day: 5/28/2011 Time: 10:35 AM
'
'    Day: 12/25/1979 Time: 14:30:00
'    Day: 12/25/1979 Time: 2:30 PM


