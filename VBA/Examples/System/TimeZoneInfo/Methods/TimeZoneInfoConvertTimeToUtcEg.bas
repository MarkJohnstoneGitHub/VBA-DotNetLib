Attribute VB_Name = "TimeZoneInfoConvertTimeToUtcEg"
'@Folder "Examples.System.TimeZoneInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 24, 2023
'@LastModified January 19, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.converttimetoutc?view=netframework-4.8.1#system-timezoneinfo-converttimetoutc(system-datetime)

Option Explicit

''
' The following example illustrates the conversion of time values whose Kind property
' is DateTimeKind.Utc, DateTimeKind.Local, and DateTimeKind.Unspecified, respectively.
' It also illustrates the conversion of ambiguous and invalid times.
''
Public Sub TimeZoneInfoConvertTimeToUtc()
    Dim datNowLocal As DotNetLib.DateTime
    Set datNowLocal = DateTime.Now
    Debug.Print VBString.Format("Converting {0}, Kind {1}:", datNowLocal, DateTimeKindHelper.ToString(datNowLocal.Kind))
    Debug.Print VBString.Format("   ConvertTimeToUtc: {0}, Kind {1}", TimeZoneInfo.ConvertTimeToUtc(datNowLocal), _
                    DateTimeKindHelper.ToString(TimeZoneInfo.ConvertTimeToUtc(datNowLocal).Kind))
    Debug.Print
    
    Dim datNowUtc As DotNetLib.DateTime
    Set datNowUtc = DateTime.UtcNow
    Debug.Print VBString.Format("Converting {0}, Kind {1}", datNowUtc, DateTimeKindHelper.ToString(datNowUtc.Kind))
    Debug.Print VBString.Format("   ConvertTimeToUtc: {0}, Kind {1}", TimeZoneInfo.ConvertTimeToUtc(datNowUtc), _
                    DateTimeKindHelper.ToString(TimeZoneInfo.ConvertTimeToUtc(datNowUtc).Kind))
    Debug.Print

    Dim datNow As DotNetLib.DateTime
    Set datNow = DateTime.CreateFromDateTime(2007, 10, 26, 13, 32, 0)
    Debug.Print VBString.Format("Converting {0}, Kind {1}", datNow, DateTimeKindHelper.ToString(datNow.Kind))
    Debug.Print VBString.Format("   ConvertTimeToUtc: {0}, Kind {1}", TimeZoneInfo.ConvertTimeToUtc(datNow), _
                    DateTimeKindHelper.ToString(TimeZoneInfo.ConvertTimeToUtc(datNow).Kind))
    Debug.Print
    
    Dim datAmbiguous As DotNetLib.DateTime
    Set datAmbiguous = DateTime.CreateFromDateTime(2007, 11, 4, 1, 30, 0)
    Debug.Print VBString.Format("Converting {0}, Kind {1}, Ambiguous {2}", datAmbiguous, _
                    DateTimeKindHelper.ToString(datAmbiguous.Kind), _
                    TimeZoneInfo.Locale.IsAmbiguousTime(datAmbiguous))
    Debug.Print VBString.Format("   ConvertTimeToUtc: {0}, Kind {1}", TimeZoneInfo.ConvertTimeToUtc(datAmbiguous), _
                    DateTimeKindHelper.ToString(TimeZoneInfo.ConvertTimeToUtc(datAmbiguous).Kind))
    Debug.Print
    
    Dim datInvalid As DotNetLib.DateTime
    Set datInvalid = DateTime.CreateFromDateTime(2007, 3, 11, 2, 30, 0)
    Debug.Print VBString.Format("Converting {0}, Kind {1}, Invalid {2}", datInvalid, _
                    datInvalid.Kind, _
                    TimeZoneInfo.Locale.IsInvalidTime(datInvalid))
    On Error Resume Next
    Debug.Print VBString.Format("   ConvertTimeToUtc: {0}, Kind {1}", _
                    TimeZoneInfo.ConvertTimeToUtc(datInvalid), _
                    DateTimeKindHelper.ToString(TimeZoneInfo.ConvertTimeToUtc(datInvalid).Kind))
    If Catch(ArgumentException) Then
        Debug.Print VBString.Format("ArgumentException: Cannot convert {1} to UTC.", datInvalid)
        Debug.Print Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
    Debug.Print

    Dim datNearMax As DotNetLib.DateTime
    Set datNearMax = DateTime.CreateFromDateTime(9999, 12, 31, 22, 0, 0)
    Debug.Print VBString.Format("Converting {0}, Kind {1}", datNearMax, DateTimeKindHelper.ToString(datNearMax.Kind))
    Debug.Print VBString.Format("   ConvertTimeToUtc: {0}, Kind {1}", _
                    TimeZoneInfo.ConvertTimeToUtc(datNearMax), _
                    DateTimeKindHelper.ToString(TimeZoneInfo.ConvertTimeToUtc(datNearMax).Kind))
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


