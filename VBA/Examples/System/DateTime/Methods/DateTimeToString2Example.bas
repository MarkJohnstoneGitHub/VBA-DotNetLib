Attribute VB_Name = "DateTimeToString2Example"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 13, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.tostring?view=netframework-4.8.1#system-datetime-tostring(system-string)

Option Explicit

' The following example uses each of the standard date and time format strings
' and a selection of custom date and time format strings to display the string
' representation of a DateTime value. The thread current culture for the
' example is en-US
Public Sub DateTimeToString2()
    Dim dateValue As DotNetLib.DateTime
    Set dateValue = DateTime.CreateFromDateTime(2008, 6, 15, 21, 15, 7)
    
    ' Create an array of standard format strings.
    Dim standardFmts() As String
    standardFmts = StringArray.CreateInitialize1D("d", "D", "f", "F", "g", "G", "m", "o", _
                                    "R", "s", "t", "T", "u", "U", "y")
                                    
    ' Output date and time using each custom format string.
    Dim standardFmt As Variant
    For Each standardFmt In standardFmts
        Debug.Print VBString.Format("{0}: {1}", standardFmt, _
                           dateValue.ToString2(standardFmt))
    Next
    Debug.Print
    
    ' Create an array of some custom format strings.
    Dim customFmts() As String
    customFmts = StringArray.CreateInitialize1D("h:mm:ss.ff t", "d MMM yyyy", "HH:mm:ss.f", _
                                    "dd MMM HH:mm:ss", "\Mon\t\h\: M", "HH:mm:ss.ffffzzz")
    ' Output date and time using each custom format string.
    Dim customFmt As Variant
    For Each customFmt In customFmts
        Debug.Print VBString.Format("'{0}': {1}", customFmt, _
                           dateValue.ToString2(customFmt))
    Next
    Debug.Print
End Sub

' This example displays the following output to the console:
'       d: 6/15/2008
'       D: Sunday, June 15, 2008
'       f: Sunday, June 15, 2008 9:15 PM
'       F: Sunday, June 15, 2008 9:15:07 PM
'       g: 6/15/2008 9:15 PM
'       G: 6/15/2008 9:15:07 PM
'       m: June 15
'       o: 2008-06-15T21:15:07.0000000
'       R: Sun, 15 Jun 2008 21:15:07 GMT
'       s: 2008-06-15T21:15:07
'       t: 9:15 PM
'       T: 9:15:07 PM
'       u: 2008-06-15 21:15:07Z
'       U: Monday, June 16, 2008 4:15:07 AM
'       y: June, 2008
'
'       'h:mm:ss.ff t': 9:15:07.00 P
'       'd MMM yyyy': 15 Jun 2008
'       'HH:mm:ss.f': 21:15:07.0
'       'dd MMM HH:mm:ss': 15 Jun 21:15:07
'       '\Mon\t\h\: M': Month: 6
'       'HH:mm:ss.ffffzzz': 21:15:07.0000-07:00


