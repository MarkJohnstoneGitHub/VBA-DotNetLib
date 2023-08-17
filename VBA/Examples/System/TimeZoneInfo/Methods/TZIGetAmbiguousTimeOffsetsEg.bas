Attribute VB_Name = "TZIGetAmbiguousTimeOffsetsEg"
'@Folder "Examples.System.TimeZoneInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 24, 2023
'@LastModified August 1, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.getambiguoustimeoffsets?view=netframework-4.8.1

Option Explicit

' The following example defines a method named ShowPossibleUtcTimes that uses the
' GetAmbiguousTimeOffsets(DateTime) method to map an ambiguous time to its possible
' corresponding Coordinated Universal Time (UTC) times.
Public Sub TimeZoneInfoGetAmbiguousTimeOffsets()
    Debug.Print
    ShowPossibleUtcTimes DateTime.CreateFromDateTime(2007, 11, 4, 1, 0, 0), TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")

    Debug.Print
    ShowPossibleUtcTimes DateTime.CreateFromDateTimeKind(2007, 11, 4, 1, 0, 0, DateTimeKind.DateTimeKind_Local), TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")

    Debug.Print
    ShowPossibleUtcTimes DateTime.CreateFromDateTimeKind(2007, 11, 4, 0, 0, 0, DateTimeKind.DateTimeKind_Local), TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")
    
    Debug.Print
    ShowPossibleUtcTimes DateTime.CreateFromDateTimeKind(2007, 11, 4, 1, 0, 0, DateTimeKind.DateTimeKind_Unspecified), TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")
    
    Debug.Print
    ShowPossibleUtcTimes DateTime.CreateFromDateTimeKind(2007, 11, 4, 7, 0, 0, DateTimeKind.DateTimeKind_Utc), TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")
End Sub

Private Sub ShowPossibleUtcTimes(ByVal ambiguousTime As IDateTime, ByVal timeZone As ITimeZoneInfo)
    Dim pvtAmbiguousTime As IDateTime
    Set pvtAmbiguousTime = ambiguousTime
    
    ' Determine if time is ambiguous in target time zone
    If (Not timeZone.IsAmbiguousTime(pvtAmbiguousTime)) Then
        Debug.Print pvtAmbiguousTime.ToString() & " is not ambiguous in time zone " & _
                    timeZone.DisplayName & "."
    Else
        ' Display time and its time zone (local, UTC, or indicated by timeZone argument)
        Dim originalTimeZoneName As String
        If (pvtAmbiguousTime.Kind = DateTimeKind.DateTimeKind_Utc) Then
            originalTimeZoneName = "UTC"
        ElseIf (pvtAmbiguousTime.Kind = DateTimeKind.DateTimeKind_Local) Then
            originalTimeZoneName = "local time"
        Else
            originalTimeZoneName = timeZone.DisplayName
        End If
        
        Debug.Print pvtAmbiguousTime.ToString() & " " & _
                    originalTimeZoneName & " maps to the following possible times:"
                    
        ' Get ambiguous offsets
        Dim offsets() As ITimeSpan
        offsets = timeZone.GetAmbiguousTimeOffsets(pvtAmbiguousTime)
        ' Handle times not in time zone of timeZone argument
        ' Local time where timeZone is not local zone
        
        If ((pvtAmbiguousTime.Kind = DateTimeKind.DateTimeKind_Local) And Not timeZone.Equals(TimeZoneInfo.Locale)) Then
        
            Set pvtAmbiguousTime = TimeZoneInfo.ConvertTime3(pvtAmbiguousTime, TimeZoneInfo.Locale, timeZone)
            ' UTC time where timeZone is not UTC zone
        ElseIf ((pvtAmbiguousTime.Kind = DateTimeKind.DateTimeKind_Utc) And Not timeZone.Equals(TimeZoneInfo.Utc)) Then
            Set pvtAmbiguousTime = TimeZoneInfo.ConvertTime3(pvtAmbiguousTime, TimeZoneInfo.Utc, timeZone)
        End If

        ' Display each offset and its mapping to UTC
        Dim varOffset As Variant
        For Each varOffset In offsets
            Dim tzOffset As ITimeSpan
            Set tzOffset = varOffset
            If (tzOffset.Equals(timeZone.BaseUtcOffset)) Then
                Debug.Print "If " & pvtAmbiguousTime.ToString() & _
                            " is " & timeZone.StandardName & _
                            ", " & DateTime.Subtraction2(pvtAmbiguousTime, tzOffset).ToString() & " UTC"
            Else
                Debug.Print "If " & pvtAmbiguousTime.ToString() & _
                            " is " & timeZone.DaylightName & _
                            ", " & DateTime.Subtraction2(pvtAmbiguousTime, tzOffset).ToString() & " UTC"
            End If
        Next
    End If
End Sub

'
' This example produces the following output if run in the Pacific time zone:
'
'    11/4/2007 1:00:00 AM (GMT-06:00) Central Time (US & Canada) maps to the following possible times:
'    If 11/4/2007 1:00:00 AM is Central Standard Time, 11/4/2007 7:00:00 AM UTC
'    If 11/4/2007 1:00:00 AM is Central Daylight Time, 11/4/2007 6:00:00 AM UTC
'
'    11/4/2007 1:00:00 AM Pacific Standard Time is not ambiguous in time zone (GMT-06:00) Central Time (US & Canada).
'
'    11/4/2007 12:00:00 AM local time maps to the following possible times:
'    If 11/4/2007 1:00:00 AM is Central Standard Time, 11/4/2007 7:00:00 AM UTC
'    If 11/4/2007 1:00:00 AM is Central Daylight Time, 11/4/2007 6:00:00 AM UTC
'
'    11/4/2007 1:00:00 AM (GMT-06:00) Central Time (US & Canada) maps to the following possible times:
'    If 11/4/2007 1:00:00 AM is Central Standard Time, 11/4/2007 7:00:00 AM UTC
'    If 11/4/2007 1:00:00 AM is Central Daylight Time, 11/4/2007 6:00:00 AM UTC
'
'    11/4/2007 7:00:00 AM UTC maps to the following possible times:
'    If 11/4/2007 1:00:00 AM is Central Standard Time, 11/4/2007 7:00:00 AM UTC
'    If 11/4/2007 1:00:00 AM is Central Daylight Time, 11/4/2007 6:00:00 AM UTC
'


