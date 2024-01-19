Attribute VB_Name = "DateTimeTryParse2Example"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 25, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.tryparse?view=netframework-4.8.1#system-datetime-tryparse(system-string-system-iformatprovider-system-globalization-datetimestyles-system-datetime@)

Option Explicit

''
' The following example illustrates the
' DateTime.TryParse(String, IFormatProvider, DateTimeStyles, DateTime) method.
''
Public Sub DateTimeTryParse2()
    Dim dateString As String
    Dim culture As DotNetLib.CultureInfo
    Dim styles As mscorlib.DateTimeStyles
    Dim dateResult As DotNetLib.DateTime
    
    ' Parse a date and time with no styles.
    dateString = "03/01/2009 10:00 AM"
    Set culture = CultureInfo.CreateSpecificCulture("en-US")
    styles = DateTimeStyles.DateTimeStyles_None
    If (DateTime.TryParse2(dateString, culture, styles, dateResult)) Then
        Debug.Print VBString.Format("{0} converted to {1} {2}.", _
                           dateString, dateResult, DateTimeKindHelper.ToString(dateResult.Kind))
    Else
        Debug.Print VBString.Format("Unable to convert {0} to a date and time.", _
                           dateString)
    End If

    ' Parse the same date and time with the AssumeLocal style.
    styles = DateTimeStyles.DateTimeStyles_AssumeLocal
    If (DateTime.TryParse2(dateString, culture, styles, dateResult)) Then
        Debug.Print VBString.Format("{0} converted to {1} {2}.", _
                           dateString, dateResult, DateTimeKindHelper.ToString(dateResult.Kind))
    Else
        Debug.Print "Unable to convert "; dateString; " to a date and time."
    End If
    
    ' Parse a date and time that is assumed to be local.
    ' This time is five hours behind UTC. The local system's time zone is
    ' eight hours behind UTC.
    dateString = "2009/03/01T10:00:00-5:00"
    styles = DateTimeStyles.DateTimeStyles_AssumeLocal
    If (DateTime.TryParse2(dateString, culture, styles, dateResult)) Then
        Debug.Print VBString.Format("{0} converted to {1} {2}.", _
                           dateString, dateResult, DateTimeKindHelper.ToString(dateResult.Kind))
    Else
        Debug.Print VBString.Format("Unable to convert {0} to a date and time.", _
                           dateString)
    End If

    ' Attempt to convert a string in improper ISO 8601 format.
    dateString = "03/01/2009T10:00:00-5:00"
    If (DateTime.TryParse2(dateString, culture, styles, dateResult)) Then
        Debug.Print VBString.Format("{0} converted to {1} {2}.", _
                           dateString, dateResult, DateTimeKindHelper.ToString(dateResult.Kind))
    Else
        Debug.Print VBString.Format("Unable to convert {0} to a date and time.", _
                           dateString)
    End If

    ' Assume a date and time string formatted for the fr-FR culture is the local
    ' time and convert it to UTC.
    dateString = "2008-03-01 10:00"
    Set culture = CultureInfo.CreateSpecificCulture("fr-FR")
    styles = DateTimeStyles.DateTimeStyles_AdjustToUniversal Or DateTimeStyles.DateTimeStyles_AssumeLocal
    If (DateTime.TryParse2(dateString, culture, styles, dateResult)) Then
        Debug.Print VBString.Format("{0} converted to {1} {2}.", _
                           dateString, dateResult, DateTimeKindHelper.ToString(dateResult.Kind))
    Else
        Debug.Print VBString.Format("Unable to convert {0} to a date and time.", _
                           dateString)
    End If
End Sub

' The example displays the following output to the console:
'       03/01/2009 10:00 AM converted to 3/1/2009 10:00:00 AM Unspecified.
'       03/01/2009 10:00 AM converted to 3/1/2009 10:00:00 AM Local.
'       2009/03/01T10:00:00-5:00 converted to 3/1/2009 7:00:00 AM Local.
'       Unable to convert 03/01/2009T10:00:00-5:00 to a date and time.
'       2008-03-01 10:00 converted to 3/1/2008 6:00:00 PM Utc.


