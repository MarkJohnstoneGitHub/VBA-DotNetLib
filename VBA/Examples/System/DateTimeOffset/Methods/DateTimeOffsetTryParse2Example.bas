Attribute VB_Name = "DateTimeOffsetTryParse2Example"
'@Folder("Examples.System.DateTimeOffset.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 27, 2023
'@LastModified August 27, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.tryparse?view=netframework-4.8.1#system-datetimeoffset-tryparse(system-string-system-iformatprovider-system-globalization-datetimestyles-system-datetimeoffset@)

Option Explicit

' The following example calls the TryParse(String, IFormatProvider, DateTimeStyles, DateTimeOffset)
' method with a variety of DateTimeStyles values to parse some strings with various date and time formats.
Public Sub DateTimeOffsetTryParse2()
    Dim dateString As String
    Dim parsedDate As DotNetLib.DateTimeOffset
    
    dateString = "05/01/2008 6:00:00"
    ' Assume time is local
    If (DateTimeOffset.TryParse2(dateString, Nothing, _
                            DateTimeStyles.DateTimeStyles_AssumeLocal, _
                            parsedDate)) Then
        Debug.Print "'"; dateString; "'"; " was converted to "; parsedDate.ToString(); "."
    Else
        Debug.Print "Unable to parse "; "'"; dateString; "'"
    End If
    
    ' Assume time is UTC
    If (DateTimeOffset.TryParse2(dateString, Nothing, _
                            DateTimeStyles.DateTimeStyles_AssumeUniversal, _
                            parsedDate)) Then
        Debug.Print "'"; dateString; "'"; " was converted to "; parsedDate.ToString(); "."
    Else
        Debug.Print "Unable to parse "; "'"; dateString; "'"
    End If
                            
    ' Parse and convert to UTC
    dateString = "05/01/2008 6:00:00AM +5:00"
    If (DateTimeOffset.TryParse2(dateString, Nothing, _
                            DateTimeStyles.DateTimeStyles_AdjustToUniversal, _
                            parsedDate)) Then
        Debug.Print "'"; dateString; "'"; " was converted to "; parsedDate.ToString(); "."
    Else
        Debug.Print "Unable to parse "; "'"; dateString; "'"
    End If
End Sub

' The example displays the following output to the console:
'    '05/01/2008 6:00:00' was converted to 5/1/2008 6:00:00 AM -07:00.
'    '05/01/2008 6:00:00' was converted to 5/1/2008 6:00:00 AM +00:00.
'    '05/01/2008 6:00:00AM +5:00' was converted to 5/1/2008 1:00:00 AM +00:00.
