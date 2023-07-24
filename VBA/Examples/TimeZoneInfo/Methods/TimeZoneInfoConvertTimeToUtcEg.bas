Attribute VB_Name = "TimeZoneInfoConvertTimeToUtcEg"
'@Folder("VBADotNetLib.Examples.TimeZoneInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 24, 2023
'@LastModified July 24, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.converttimetoutc?view=netframework-4.8.1#system-timezoneinfo-converttimetoutc(system-datetime)

Option Explicit

' The following example illustrates the conversion of time values whose Kind property
' is DateTimeKind.Utc, DateTimeKind.Local, and DateTimeKind.Unspecified, respectively.
' It also illustrates the conversion of ambiguous and invalid times.
Public Sub TimeZoneInfoConvertTimeToUtc()
    Dim datNowLocal As DateTime
    Set datNowLocal = DateTime.Now
    Debug.Print "Converting " & datNowLocal.ToString() & ", Kind " & _
                DateTimeKindHelper.ToString(datNowLocal.Kind) & ":"
    Debug.Print "   ConvertTimeToUtc: " & TimeZoneInfo.ConvertTimeToUtc(datNowLocal).ToString() & ", Kind " & _
                DateTimeKindHelper.ToString(TimeZoneInfo.ConvertTimeToUtc(datNowLocal).Kind)
    Debug.Print
    
    Dim datNowUtc As DateTime
    Set datNowUtc = DateTime.UtcNow
    Debug.Print "Converting " & datNowUtc.ToString() & ", Kind " & _
                DateTimeKindHelper.ToString(datNowUtc.Kind) & ":"
    Debug.Print "   ConvertTimeToUtc: " & TimeZoneInfo.ConvertTimeToUtc(datNowUtc).ToString() & ", Kind " & _
                DateTimeKindHelper.ToString(TimeZoneInfo.ConvertTimeToUtc(datNowUtc).Kind)
    Debug.Print
    
    Dim datNow As DateTime
    Set datNow = DateTime.CreateFromDateTime(2007, 10, 26, 13, 32, 0)
    Debug.Print "Converting " & datNow.ToString() & ", Kind " & _
                DateTimeKindHelper.ToString(datNow.Kind) & ":"
    Debug.Print "   ConvertTimeToUtc: " & TimeZoneInfo.ConvertTimeToUtc(datNow).ToString() & ", Kind " & _
                DateTimeKindHelper.ToString(TimeZoneInfo.ConvertTimeToUtc(datNow).Kind)
    Debug.Print
    
    Dim datAmbiguous As DateTime
    Set datAmbiguous = DateTime.CreateFromDateTime(2007, 11, 4, 1, 30, 0)
    Debug.Print "Converting " & datAmbiguous.ToString() & ", Kind " & _
                DateTimeKindHelper.ToString(datAmbiguous.Kind) & ":" & _
                ", Ambiguous " & TimeZoneInfo.Locale.IsAmbiguousTime(datAmbiguous)
    Debug.Print "   ConvertTimeToUtc: " & TimeZoneInfo.ConvertTimeToUtc(datAmbiguous).ToString() & ", Kind " & _
                DateTimeKindHelper.ToString(TimeZoneInfo.ConvertTimeToUtc(datAmbiguous).Kind)
    Debug.Print
    
    Dim datInvalid As DateTime
    Set datInvalid = DateTime.CreateFromDateTime(2007, 3, 11, 2, 30, 0)
    Debug.Print "Converting " & datInvalid.ToString() & ", Kind " & _
                DateTimeKindHelper.ToString(datInvalid.Kind) & ":" & _
                ", Invalid " & TimeZoneInfo.Locale.IsInvalidTime(datInvalid)
                
    On Error Resume Next
    Debug.Print "   ConvertTimeToUtc: " & TimeZoneInfo.ConvertTimeToUtc(datInvalid).ToString() & ", Kind " & _
                DateTimeKindHelper.ToString(TimeZoneInfo.ConvertTimeToUtc(datInvalid).Kind)
    If Catch(ArgumentException) Then
        Debug.Print "ArgumentException: Cannot convert " & datInvalid.ToString() & " to UTC."
        Debug.Print Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
    Debug.Print

    Dim datNearMax As DateTime
    Set datNearMax = DateTime.CreateFromDateTime(9999, 12, 31, 22, 0, 0)
    
    Debug.Print "Converting " & datNearMax.ToString() & ", Kind " & _
                DateTimeKindHelper.ToString(datNearMax.Kind)
    Debug.Print "   ConvertTimeToUtc: " & TimeZoneInfo.ConvertTimeToUtc(datNearMax).ToString() & ", Kind " & _
                DateTimeKindHelper.ToString(TimeZoneInfo.ConvertTimeToUtc(datNearMax).Kind)
    Debug.Print
End Sub

' This example produces the following output if the local time zone
' is Pacific Standard Time:
'
'    Converting 8/31/2007 2:26:28 PM, Kind Local:
'       ConvertTimeToUtc: 8/31/2007 9:26:28 PM, Kind Utc
'
'    Converting 8/31/2007 9:26:28 PM, Kind Utc
'       ConvertTimeToUtc: 8/31/2007 9:26:28 PM, Kind Utc
'
'    Converting 10/26/2007 1:32:00 PM, Kind Unspecified
'       ConvertTimeToUtc: 10/26/2007 8:32:00 PM, Kind Utc
'
'    Converting 11/4/2007 1:30:00 AM, Kind Unspecified, Ambiguous True
'       ConvertTimeToUtc: 11/4/2007 9:30:00 AM, Kind Utc
'
'    Converting 3/11/2007 2:30:00 AM, Kind Unspecified, Invalid True
'       ArgumentException: Cannot convert 3/11/2007 2:30:00 AM to UTC.
'
'    Converting 12/31/9999 10:00:00 PM, Kind Unspecified
'       ConvertTimeToUtc: 12/31/9999 11:59:59 PM, Kind Utc
'
