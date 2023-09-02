Attribute VB_Name = "DateTimeTryParseExactExample"
'@Folder("Examples.System.DateTime.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 15, 2023
'@LastModified September 2, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.tryparseexact?view=netframework-4.8.1#system-datetime-tryparseexact(system-string-system-string-system-iformatprovider-system-globalization-datetimestyles-system-datetime@)

Option Explicit

'@Description("The following example demonstrates the DateTime.TryParseExact(String, String, IFormatProvider, DateTimeStyles, DateTime) method.")
' Note that the string " 5/01/2009 8:30 AM" cannot be parsed successfully when
' the styles parameter equals DateTimeStyles.None because leading spaces are
' not allowed by format. Additionally, the string "5/01/2009 09:00" cannot be
' parsed successfully with a format of "MM/dd/yyyyhh:mm" because the date string
' does not precede the month number with a leading zero, as format requires.
Public Sub DateTimeTryParseExact()
Attribute DateTimeTryParseExact.VB_Description = "The following example demonstrates the DateTime.TryParseExact(String, String, IFormatProvider, DateTimeStyles, DateTime) method."
    Dim enUS  As DotNetLib.CultureInfo
    Set enUS = CultureInfo.CreateFromName("en-US")
    Dim dateValue As DotNetLib.DateTime
    
    ' Parse date with no style flags.
    Dim dateString As String
    dateString = " 5/01/2009 8:30 AM"
    If (DateTime.TryParseExact(dateString, "g", enUS, DateTimeStyles.DateTimeStyles_None, dateValue)) Then
        Debug.Print "Converted '"; dateString; "' to "; dateValue.ToString(); _
                    " ("; DateTimeKindHelper.ToString(dateValue.Kind); ")."
    Else
        Debug.Print "'"; dateString; "' is not in an acceptable format."
    End If
    
    ' Allow a leading space in the date string.
    If (DateTime.TryParseExact(dateString, "g", enUS, DateTimeStyles.DateTimeStyles_AllowLeadingWhite, dateValue)) Then
        Debug.Print "Converted '"; dateString; "' to "; dateValue.ToString(); _
                    " ("; DateTimeKindHelper.ToString(dateValue.Kind); ")."
    Else
        Debug.Print "'"; dateString; "' is not in an acceptable format."
    End If

    ' Use custom formats with M and MM.
    dateString = "5/01/2009 09:00"
    If (DateTime.TryParseExact(dateString, "M/dd/yyyy hh:mm", enUS, DateTimeStyles.DateTimeStyles_None, dateValue)) Then
        Debug.Print "Converted '"; dateString; "' to "; dateValue.ToString(); _
                    " ("; DateTimeKindHelper.ToString(dateValue.Kind); ")."
    Else
        Debug.Print "'"; dateString; "' is not in an acceptable format."
    End If

    ' Allow a leading space in the date string.
    If (DateTime.TryParseExact(dateString, "MM/dd/yyyy hh:mm", enUS, DateTimeStyles.DateTimeStyles_None, dateValue)) Then
        Debug.Print "Converted '"; dateString; "' to "; dateValue.ToString(); _
                    " ("; DateTimeKindHelper.ToString(dateValue.Kind); ")."
    Else
        Debug.Print "'"; dateString; "' is not in an acceptable format."
    End If

    ' Parse a string with time zone information.
    dateString = "05/01/2009 01:30:42 PM -05:00"
    If (DateTime.TryParseExact(dateString, "MM/dd/yyyy hh:mm:ss tt zzz", enUS, DateTimeStyles.DateTimeStyles_None, dateValue)) Then
        Debug.Print "Converted '"; dateString; "' to "; dateValue.ToString(); _
                    " ("; DateTimeKindHelper.ToString(dateValue.Kind); ")."
    Else
        Debug.Print "'"; dateString; "' is not in an acceptable format."
    End If

    ' Allow a leading space in the date string.
    If (DateTime.TryParseExact(dateString, "MM/dd/yyyy hh:mm:ss tt zzz", enUS, DateTimeStyles.DateTimeStyles_AdjustToUniversal, dateValue)) Then
        Debug.Print "Converted '"; dateString; "' to "; dateValue.ToString(); _
                    " ("; DateTimeKindHelper.ToString(dateValue.Kind); ")."
    Else
        Debug.Print "'"; dateString; "' is not in an acceptable format."
    End If
    
    ' Parse a string representing UTC.
    dateString = "2008-06-11T16:11:20.0904778Z"
    If (DateTime.TryParseExact(dateString, "o", CultureInfo.InvariantCulture, DateTimeStyles.DateTimeStyles_None, dateValue)) Then
        Debug.Print "Converted '"; dateString; "' to "; dateValue.ToString(); _
                    " ("; DateTimeKindHelper.ToString(dateValue.Kind); ")."
    Else
        Debug.Print "'"; dateString; "' is not in an acceptable format."
    End If
    
    If (DateTime.TryParseExact(dateString, "o", CultureInfo.InvariantCulture, DateTimeStyles.DateTimeStyles_RoundTripKind, dateValue)) Then
        Debug.Print "Converted '"; dateString; "' to "; dateValue.ToString(); _
                    " ("; DateTimeKindHelper.ToString(dateValue.Kind); ")."
    Else
        Debug.Print "'"; dateString; "' is not in an acceptable format."
    End If

End Sub

' The example displays the following output:
'    ' 5/01/2009 8:30 AM' is not in an acceptable format.
'    Converted ' 5/01/2009 8:30 AM' to 5/1/2009 8:30:00 AM (Unspecified).
'    Converted '5/01/2009 09:00' to 5/1/2009 9:00:00 AM (Unspecified).
'    '5/01/2009 09:00' is not in an acceptable format.
'    Converted '05/01/2009 01:30:42 PM -05:00' to 5/1/2009 11:30:42 AM (Local).
'    Converted '05/01/2009 01:30:42 PM -05:00' to 5/1/2009 6:30:42 PM (Utc).
'    Converted '2008-06-11T16:11:20.0904778Z' to 6/11/2008 9:11:20 AM (Local).
'    Converted '2008-06-11T16:11:20.0904778Z' to 6/11/2008 4:11:20 PM (Utc).
