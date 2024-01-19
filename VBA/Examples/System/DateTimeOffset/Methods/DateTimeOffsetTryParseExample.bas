Attribute VB_Name = "DateTimeOffsetTryParseExample"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.tryparse?view=netframework-4.8.1#system-datetimeoffset-tryparse(system-string-system-datetimeoffset@)

Option Explicit

''
' The following example calls the TryParse(String, DateTimeOffset) method to
' parse several strings with various date and time formats.
''
Public Sub DateTimeOffsetTryParse()
    Dim parsedDate As DotNetLib.DateTimeOffset
    Dim dateString As String
    
    '// String with date only
    dateString = "05/01/2008"
    If (DateTimeOffset.TryParse(dateString, parsedDate)) Then
        Debug.Print VBString.Format("{0} was converted to {1}.", _
                                    dateString, parsedDate)
    End If
    
    ' String with time only
    dateString = "11:36 PM"
    If (DateTimeOffset.TryParse(dateString, parsedDate)) Then
        Debug.Print VBString.Format("{0} was converted to {1}.", _
                                    dateString, parsedDate)
    End If

    ' String with date and offset
    dateString = "05/01/2008 +7:00"
    If (DateTimeOffset.TryParse(dateString, parsedDate)) Then
        Debug.Print VBString.Format("{0} was converted to {1}.", _
                                    dateString, parsedDate)
    End If

    '// String with day abbreviation
    dateString = "Thu May 01, 2008"
    If (DateTimeOffset.TryParse(dateString, parsedDate)) Then
        Debug.Print VBString.Format("{0} was converted to {1}.", _
                                    dateString, parsedDate)
    End If
    
    ' String with date, time with AM/PM designator, and offset
    dateString = "5/1/2008 10:00 AM -07:00"
    If (DateTimeOffset.TryParse(dateString, parsedDate)) Then
        Debug.Print VBString.Format("{0} was converted to {1}.", _
                                    dateString, parsedDate)
    End If
End Sub

' if (run on 3/29/07, the example displays the following output:
'    05/01/2008 was converted to 5/1/2008 12:00:00 AM -07:00.
'    11:36 PM was converted to 3/29/2007 11:36:00 PM -07:00.
'    05/01/2008 +7:00 was converted to 5/1/2008 12:00:00 AM +07:00.
'    Thu May 01, 2008 was converted to 5/1/2008 12:00:00 AM -07:00.
'    5/1/2008 10:00 AM -07:00 was converted to 5/1/2008 10:00:00 AM -07:00.

