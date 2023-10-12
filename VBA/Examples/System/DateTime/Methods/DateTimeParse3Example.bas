Attribute VB_Name = "DateTimeParse3Example"
'@Folder("Examples.System.DateTime.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 18, 2023
'@LastModified August 18, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.parse?view=netframework-4.8.1#system-datetime-parse(system-string-system-iformatprovider-system-globalization-datetimestyles)

Option Explicit

' The following example demonstrates the Parse(String, IFormatProvider, DateTimeStyles)
' method and displays the value of the Kind property of the resulting DateTime values.
Public Sub DateTimeParse3()
    Dim dateString As String
    Dim culture As DotNetLib.CultureInfo
    Dim styles As mscorlib.DateTimeStyles
    Dim Result As DotNetLib.DateTime
    
    ' Parse a date and time with no styles.
    dateString = "03/01/2009 10:00 AM"
    Set culture = CultureInfo.CreateSpecificCulture("en-US")
    styles = DateTimeStyles.DateTimeStyles_None
    
    On Error Resume Next
    Set Result = DateTime.Parse3(dateString, culture, styles)
    If Try Then
        Debug.Print dateString; " converted to "; Result.ToString(); " "; DateTimeKindHelper.ToString(Result.Kind)
    ElseIf Catch(FormatException) Then
        Debug.Print "Unable to convert "; dateString; " to a date and time."
    End If
    On Error GoTo 0 'reset error handling
    
    ' Parse the same date and time with the AssumeLocal style.
    styles = DateTimeStyles.DateTimeStyles_AssumeLocal
    Set Result = DateTime.Parse3(dateString, culture, styles)
    If Try Then
        Debug.Print dateString; " converted to "; Result.ToString(); " "; DateTimeKindHelper.ToString(Result.Kind)
    ElseIf Catch(FormatException) Then
        Debug.Print "Unable to convert "; dateString; " to a date and time."
    End If
    On Error GoTo 0 'reset error handling
    
    ' Parse a date and time that is assumed to be local.
    ' This time is five hours behind UTC. The local system's time zone is
    ' eight hours behind UTC.
    dateString = "2009/03/01T10:00:00-5:00"
    styles = DateTimeStyles.DateTimeStyles_AssumeLocal
    On Error Resume Next
    Set Result = DateTime.Parse3(dateString, culture, styles)
    If Try Then
        Debug.Print dateString; " converted to "; Result.ToString(); " "; DateTimeKindHelper.ToString(Result.Kind)
    ElseIf Catch(FormatException) Then
        Debug.Print "Unable to convert "; dateString; " to a date and time."
    End If
    On Error GoTo 0 'reset error handling

    ' Attempt to convert a string in improper ISO 8601 format.
    dateString = "03/01/2009T10:00:00-5:00"
    On Error Resume Next
    Set Result = DateTime.Parse3(dateString, culture, styles)
    If Try Then
        Debug.Print dateString; " converted to "; Result.ToString(); " "; DateTimeKindHelper.ToString(Result.Kind)
    ElseIf Catch(FormatException) Then
        Debug.Print "Unable to convert "; dateString; " to a date and time."
    End If
    On Error GoTo 0 'reset error handling
    
    ' Assume a date and time string formatted for the fr-FR culture is the local
    ' time and convert it to UTC.
    dateString = "2008-03-01 10:00"
    Set culture = CultureInfo.CreateSpecificCulture("fr-FR")
    styles = DateTimeStyles.DateTimeStyles_AdjustToUniversal Or DateTimeStyles.DateTimeStyles_AssumeLocal
    On Error Resume Next
    Set Result = DateTime.Parse3(dateString, culture, styles)
    If Try Then
        Debug.Print dateString; " converted to "; Result.ToString(); " "; DateTimeKindHelper.ToString(Result.Kind)
    ElseIf Catch(FormatException) Then
        Debug.Print "Unable to convert "; dateString; " to a date and time."
    End If
    On Error GoTo 0 'reset error handling

End Sub

' The example displays the following output to the console:
'       03/01/2009 10:00 AM converted to 3/1/2009 10:00:00 AM Unspecified.
'       03/01/2009 10:00 AM converted to 3/1/2009 10:00:00 AM Local.
'       2009/03/01T10:00:00-5:00 converted to 3/1/2009 7:00:00 AM Local.
'       Unable to convert 03/01/2009T10:00:00-5:00 to a date and time.
'       2008-03-01 10:00 converted to 3/1/2008 6:00:00 PM Utc.
