Attribute VB_Name = "DateTimeOffsetParseExactExample"
'@Folder("Examples.System.DateTimeOffset.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 25, 2023
'@LastModified August 25, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.parseexact?view=netframework-4.8.1#system-datetimeoffset-parseexact(system-string-system-string-system-iformatprovider)

Option Explicit

' The following example uses the DateTimeOffset.ParseExact(String, String, IFormatProvider)
' method with standard and custom format specifiers and the invariant culture to parse
' several date and time strings.
Public Sub DateTimeOffsetParseExact()
    Dim dateString As String
    Dim pvtFormat As String
    Dim Result As DotNetLib.DateTimeOffset
    Dim provider As DotNetLib.CultureInfo
    Set provider = CultureInfo.InvariantCulture
    
    ' Parse date-only value with invariant culture.
    dateString = "06/15/2008"
    pvtFormat = "d"
    On Error Resume Next
    Set Result = DateTimeOffset.ParseExact(dateString, pvtFormat, provider)
    If Try Then
        Debug.Print dateString; " converts to "; Result.ToString()
    ElseIf Catch(FormatException) Then
       Debug.Print dateString; " is not in the correct format."
    End If
    On Error GoTo 0 'reset error handling
    
    ' Parse date-only value without leading zero in month using "d" format.
    ' Should throw a FormatException because standard short date pattern of
    ' invariant culture requires two-digit month.
    dateString = "6/15/2008"
    On Error Resume Next
    Set Result = DateTimeOffset.ParseExact(dateString, pvtFormat, provider)
    If Try Then
        Debug.Print dateString; " converts to "; Result.ToString()
    ElseIf Catch(FormatException) Then
       Debug.Print dateString; " is not in the correct format."
    End If
    On Error GoTo 0 'reset error handling
    
    ' Parse date and time with custom specifier.
    dateString = "Sun 15 Jun 2008 8:30 AM -06:00"
    pvtFormat = "ddd dd MMM yyyy h:mm tt zzz"
    On Error Resume Next
    Set Result = DateTimeOffset.ParseExact(dateString, pvtFormat, provider)
    If Try Then
        Debug.Print dateString; " converts to "; Result.ToString()
    ElseIf Catch(FormatException) Then
       Debug.Print dateString; " is not in the correct format."
    End If
    On Error GoTo 0 'reset error handling
    
    ' Parse date and time with offset without offset//s minutes.
    ' Should throw a FormatException because "zzz" specifier requires leading
    ' zero in hours.
    dateString = "Sun 15 Jun 2008 8:30 AM -06"
    On Error Resume Next
    Set Result = DateTimeOffset.ParseExact(dateString, pvtFormat, provider)
    If Try Then
        Debug.Print dateString; " converts to "; Result.ToString()
    ElseIf Catch(FormatException) Then
       Debug.Print dateString; " is not in the correct format."
    End If
    On Error GoTo 0 'reset error handling
End Sub

' The example displays the following output:
'    06/15/2008 converts to 6/15/2008 12:00:00 AM -07:00.
'    6/15/2008 is not in the correct format.
'    Sun 15 Jun 2008 8:30 AM -06:00 converts to 6/15/2008 8:30:00 AM -06:00.
'    Sun 15 Jun 2008 8:30 AM -06 is not in the correct format.
