Attribute VB_Name = "DateTimeOffsetParseExact2Eg"
'@Folder("Examples.System.DateTimeOffset.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 25, 2023
'@LastModified August 25, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.parseexact?view=netframework-4.8.1#system-datetimeoffset-parseexact(system-string-system-string-system-iformatprovider-system-globalization-datetimestyles)

Option Explicit

' The following example uses the DateTimeOffset.ParseExact(String, String, IFormatProvider, DateTimeStyles)
' method with standard and custom format specifiers, the invariant culture, and various DateTimeStyles
' values to parse several date and time strings.
Public Sub DateTimeOffsetParseExact2()
    Dim dateString As String
    Dim pvtFormat As String
    Dim result As DotNetLib.DateTimeOffset
    Dim provider As DotNetLib.CultureInfo
    Set provider = CultureInfo.InvariantCulture

    ' Parse date-only value with invariant culture and assume time is UTC.
    dateString = "06/15/2008"
    pvtFormat = "d"
    On Error Resume Next
    Set result = DateTimeOffset.ParseExact2(dateString, pvtFormat, provider, _
                                    DateTimeStyles.DateTimeStyles_AssumeUniversal)
    If Try Then
        Debug.Print "'"; dateString; "'"; " converts to "; result.ToString()
    ElseIf Catch(FormatException) Then
       Debug.Print "'"; dateString; "'"; " is not in the correct format."
    End If
    On Error GoTo 0 'reset error handling
    
    ' Parse date-only value with leading white space.
    ' Should throw a FormatException because only trailing white space is
    ' specified in method call.
    dateString = " 06/15/2008"
    On Error Resume Next
    Set result = DateTimeOffset.ParseExact2(dateString, pvtFormat, provider, _
                                    DateTimeStyles.DateTimeStyles_AllowTrailingWhite)
    If Try Then
        Debug.Print "'"; dateString; "'"; " converts to "; result.ToString()
    ElseIf Catch(FormatException) Then
       Debug.Print "'"; dateString; "'"; " is not in the correct format."
    End If
    On Error GoTo 0 'reset error handling
    
    ' Parse date and time value, and allow all white space.
    dateString = " 06/15/   2008  15:15    -05:00"
    pvtFormat = "MM/dd/yyyy H:mm zzz"
    On Error Resume Next
    Set result = DateTimeOffset.ParseExact2(dateString, pvtFormat, provider, _
                                    DateTimeStyles.DateTimeStyles_AllowWhiteSpaces)
    If Try Then
        Debug.Print "'"; dateString; "'"; " converts to "; result.ToString()
    ElseIf Catch(FormatException) Then
       Debug.Print "'"; dateString; "'"; " is not in the correct format."
    End If
    On Error GoTo 0 'reset error handling
    
    ' Parse date and time and convert to UTC.
    dateString = "  06/15/2008 15:15:30 -05:00"
    pvtFormat = "MM/dd/yyyy H:mm:ss zzz"
    On Error Resume Next
    Set result = DateTimeOffset.ParseExact2(dateString, pvtFormat, provider, _
                                    DateTimeStyles.DateTimeStyles_AllowWhiteSpaces Or DateTimeStyles.DateTimeStyles_AdjustToUniversal)
    If Try Then
        Debug.Print "'"; dateString; "'"; " converts to "; result.ToString()
    ElseIf Catch(FormatException) Then
       Debug.Print "'"; dateString; "'"; " is not in the correct format."
    End If
    On Error GoTo 0 'reset error handling
End Sub

' The example displays the following output:
'    '06/15/2008' converts to 6/15/2008 12:00:00 AM +00:00.
'    ' 06/15/2008' is not in the correct format.
'    ' 06/15/   2008  15:15    -05:00' converts to 6/15/2008 3:15:00 PM -05:00.
'    ' 06/15/2008 15:15:30 -05:00' converts to 6/15/2008 8:15:30 PM +00:00.
