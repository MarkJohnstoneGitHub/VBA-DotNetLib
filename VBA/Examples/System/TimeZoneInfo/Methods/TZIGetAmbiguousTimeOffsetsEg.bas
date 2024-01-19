Attribute VB_Name = "TZIGetAmbiguousTimeOffsetsEg"
'@Folder "Examples.System.TimeZoneInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 24, 2023
'@LastModified January 19, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.getambiguoustimeoffsets?view=netframework-4.8.1

Option Explicit

''
' The following example defines a method named ShowPossibleUtcTimes that uses
' the GetAmbiguousTimeOffsets(DateTime) method to map an ambiguous time to its
' possible corresponding Coordinated Universal Time (UTC) times.
''
Public Sub TimeZoneInfoGetAmbiguousTimeOffsets()
    Debug.Print
    ShowPossibleUtcTimes DateTime.CreateFromDateTime(2007, 11, 4, 1, 0, 0), _
                        TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")

    Debug.Print
    ShowPossibleUtcTimes DateTime.CreateFromDateTimeKind(2007, 11, 4, 1, 0, 0, DateTimeKind.DateTimeKind_Local), _
                        TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")

    Debug.Print
    ShowPossibleUtcTimes DateTime.CreateFromDateTimeKind(2007, 11, 4, 0, 0, 0, DateTimeKind.DateTimeKind_Local), _
                        TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")
    
    Debug.Print
    ShowPossibleUtcTimes DateTime.CreateFromDateTimeKind(2007, 11, 4, 1, 0, 0, DateTimeKind.DateTimeKind_Unspecified), _
                        TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")
    
    Debug.Print
    ShowPossibleUtcTimes DateTime.CreateFromDateTimeKind(2007, 11, 4, 7, 0, 0, DateTimeKind.DateTimeKind_Utc), _
                        TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time")
End Sub

Private Sub ShowPossibleUtcTimes(ByVal ambiguousTime As DotNetLib.DateTime, ByVal timeZone As DotNetLib.TimeZoneInfo)
    Dim pvtAmbiguousTime As DotNetLib.DateTime
    Set pvtAmbiguousTime = ambiguousTime
    
    ' Determine if time is ambiguous in target time zone
    If (Not timeZone.IsAmbiguousTime(pvtAmbiguousTime)) Then
        Debug.Print VBString.Format("{0} is not ambiguous in time zone {1}.", _
                        pvtAmbiguousTime, _
                        timeZone.DisplayName)
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
        
        Debug.Print VBString.Format("{0} {1} maps to the following possible times:", _
                  pvtAmbiguousTime, originalTimeZoneName)
                    
        ' Get ambiguous offsets
        Dim offsets() As DotNetLib.TimeSpan
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
            Dim tzOffset As DotNetLib.TimeSpan
            Set tzOffset = varOffset
            If (tzOffset.Equals(timeZone.BaseUtcOffset)) Then
                Debug.Print VBString.Format("If {0} is {1}, {2} UTC", _
                                pvtAmbiguousTime, _
                                timeZone.StandardName, _
                                DateTime.Subtraction2(pvtAmbiguousTime, tzOffset))
            Else
                Debug.Print VBString.Format("If {0} is {1}, {2} UTC", _
                                pvtAmbiguousTime, _
                                timeZone.DaylightName, _
                                DateTime.Subtraction2(pvtAmbiguousTime, tzOffset))
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


