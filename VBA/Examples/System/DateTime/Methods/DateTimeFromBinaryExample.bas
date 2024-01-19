Attribute VB_Name = "DateTimeFromBinaryExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 11, 2023
'@LastModified January 6, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.frombinary?view=netframework-4.8.1#local-time-adjustment

Option Explicit

Public Sub DateTimeFromBinary()
    Dim localDate As DotNetLib.DateTime
    Set localDate = DateTime.CreateFromDateTimeKind(2023, 10, 1, 2, 30, 0, DateTimeKind.DateTimeKind_Local)
    
    Dim binLocal As LongLong
    binLocal = localDate.ToBinary()
    
    If (TimeZoneInfo.Locale.IsInvalidTime(localDate)) Then
        Debug.Print VBString.Format("{0} is an invalid time in the {1} zone.", _
                           localDate, _
                           TimeZoneInfo.Locale.StandardName)
    End If
    Dim localDate2 As DotNetLib.DateTime
    Set localDate2 = DateTime.FromBinary(binLocal)
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                        localDate, localDate2, localDate.Equals(localDate2))
End Sub

' The example displays the following output for the local time zone AUS Eastern Standard Time:
'    1/10/2023 2:30:00 AM is an invalid time in the AUS Eastern Standard Time
'    1/10/2023 2:30:00 AM = 1/10/2023 3:30:00 AM: False

' The example displays the following output for the local time zone US Pacific Standard Time:
'    10/1/2023 2:30:00 AM = 10/1/2023 2:30:00 AM: True


