Attribute VB_Name = "DateTimeOffsetTryParseExactEg"
'@Folder("Examples.System.DateTimeOffset.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 27, 2023
'@LastModified August 27, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.tryparseexact?view=netframework-4.8.1#system-datetimeoffset-tryparseexact(system-string-system-string-system-iformatprovider-system-globalization-datetimestyles-system-datetimeoffset@)

Option Explicit

' The following example uses the TryParseExact(String, String, IFormatProvider, DateTimeStyles, DateTimeOffset)
' method with standard and custom format specifiers, the invariant culture, and various DateTimeStyles values
' to parse several date and time strings.
Public Sub DateTimeOffsetTryParseExact()
    Dim dateString As String
    Dim pvtFormat As String
    Dim result As DotNetLib.DateTimeOffset
    Dim provider As mscorlib.IFormatProvider
    
    ' Parse date-only value with invariant culture and assume time is UTC.
    dateString = "06/15/2008"
    pvtFormat = "d"
    If (DateTimeOffset.TryParseExact(dateString, pvtFormat, provider, _
                                DateTimeStyles.DateTimeStyles_AssumeUniversal, _
                                result)) Then
        Debug.Print "'"; dateString; "'"; " converts to "; result.ToString()
    Else
        Debug.Print "'"; dateString; "'"; " is not in the correct format."
    End If
    
    
    ' Parse date-only value with leading white space.
    ' Should return False because only trailing white space is
    ' specified in method call.
    dateString = " 06/15/2008"
    If (DateTimeOffset.TryParseExact(dateString, pvtFormat, provider, _
                                    DateTimeStyles.DateTimeStyles_AllowTrailingWhite, _
                                    result)) Then
        Debug.Print "'"; dateString; "'"; " converts to "; result.ToString()
    Else
        Debug.Print "'"; dateString; "'"; " is not in the correct format."
    End If

    ' Parse date and time value, and allow all white space.
    dateString = " 06/15/   2008  15:15    -05:00"
    pvtFormat = "MM/dd/yyyy H:mm zzz"
    If (DateTimeOffset.TryParseExact(dateString, pvtFormat, provider, _
                                    DateTimeStyles.DateTimeStyles_AllowWhiteSpaces, _
                                    result)) Then
        Debug.Print "'"; dateString; "'"; " converts to "; result.ToString()
    Else
        Debug.Print "'"; dateString; "'"; " is not in the correct format."
    End If
    
    ' Parse date and time and convert to UTC.
    dateString = "  06/15/2008 15:15:30 -05:00"
    pvtFormat = "MM/dd/yyyy H:mm:ss zzz"
    If (DateTimeOffset.TryParseExact(dateString, pvtFormat, provider, _
                                DateTimeStyles.DateTimeStyles_AllowWhiteSpaces Or _
                                DateTimeStyles.DateTimeStyles_AdjustToUniversal, _
                                result)) Then
        Debug.Print "'"; dateString; "'"; " converts to "; result.ToString()
    Else
        Debug.Print "'"; dateString; "'"; " is not in the correct format."
    End If
End Sub

' The example displays the following output:
'    '06/15/2008' converts to 6/15/2008 12:00:00 AM +00:00.
'    ' 06/15/2008' is not in the correct format.
'    ' 06/15/   2008  15:15    -05:00' converts to 6/15/2008 3:15:00 PM -05:00.
'    '  06/15/2008 15:15:30 -05:00' converts to 6/15/2008 8:15:30 PM +00:00.
