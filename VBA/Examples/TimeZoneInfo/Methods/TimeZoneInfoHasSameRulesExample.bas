Attribute VB_Name = "TimeZoneInfoHasSameRulesExample"
'@Folder("VBADotNetLib.Examples.TimeZoneInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 26, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.hassamerules?view=netframework-4.8.1#examples

Option Explicit

' Typically, a number of time zones defined in the registry on Windows and the ICU Library on Linux
' and macOS have the same offset from Coordinated Universal Time (UTC) and the same adjustment rules.
' The following example displays a list of these time zones to the console.
Public Sub TimeZoneInfoHasSameRules()
    Dim timeZones As DotNetLib.IReadOnlyCollection
    Set timeZones = TimeZoneInfo.GetSystemTimeZones()
    Dim timeZoneArray() As Variant
    ReDim timeZoneArray(timeZones.Count - 1)
    timeZones.CopyTo timeZoneArray, 0

    ' Iterate array from top to bottom
    Dim ctr As Long
    For ctr = UBound(timeZoneArray) To LBound(timeZoneArray) Step -1
        ' Get next item from top
        Dim thisTimeZone As DotNetLib.ITimeZoneInfo
        Set thisTimeZone = timeZoneArray(ctr)
        Dim compareCtr As Long
        compareCtr = 0
        For compareCtr = 0 To ctr - 1
            ' Determine if time zones have the same rules
            If thisTimeZone.HasSameRules(timeZoneArray(compareCtr)) Then
                Debug.Print thisTimeZone.StandardName; " has the same rules as " & _
                            timeZoneArray(compareCtr).StandardName
            End If
        Next
    Next
End Sub

' Output:
'    West Pacific Standard Time has the same rules as E. Australia Standard Time
'    Korea Standard Time has the same rules as Tokyo Standard Time
'    Taipei Standard Time has the same rules as China Standard Time
'    Taipei Standard Time has the same rules as Malay Peninsula Standard Time
'    Malay Peninsula Standard Time has the same rules as China Standard Time
'    Sri Lanka Standard Time has the same rules as India Standard Time
'    Georgian Standard Time has the same rules as Arabian Standard Time
'    E. Africa Standard Time has the same rules as Arab Standard Time
'    FLE Standard Time has the same rules as GTB Standard Time
'    Central European Standard Time has the same rules as W. Europe Standard Time
'    Central European Standard Time has the same rules as Central Europe Standard Time
'    Central European Standard Time has the same rules as Romance Standard Time
'    Romance Standard Time has the same rules as W. Europe Standard Time
'    Romance Standard Time has the same rules as Central Europe Standard Time
'    Central Europe Standard Time has the same rules as W. Europe Standard Time
'    Greenwich Standard Time has the same rules as Coordinated Universal Time
'    Canada Central Standard Time has the same rules as Central America Standard Time
