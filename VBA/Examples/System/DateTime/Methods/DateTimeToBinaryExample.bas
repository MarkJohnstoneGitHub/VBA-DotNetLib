Attribute VB_Name = "DateTimeToBinaryExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 13, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.tobinary?view=netframework-4.8.1#local-time-adjustment

Option Explicit

'@Description("Serializes the current DateTime object to a 64-bit binary value that subsequently can be used to recreate the DateTime object.")
Public Sub DateTimeToBinary()
Attribute DateTimeToBinary.VB_Description = "Serializes the current DateTime object to a 64-bit binary value that subsequently can be used to recreate the DateTime object."
    Dim localDate As DotNetLib.DateTime
    Set localDate = DateTime.CreateFromDateTimeKind(2010, 3, 14, 2, 30, 0, DateTimeKind.DateTimeKind_Local)
    Dim binLocal As LongLong
    binLocal = localDate.ToBinary()
    If (TimeZoneInfo.Locale.IsInvalidTime(localDate)) Then
        Debug.Print VBString.Format("{0} is an invalid time in the {1} zone.", _
                           localDate, _
                           TimeZoneInfo.Locale.StandardName)
    Else
        Debug.Print VBString.Format("{0} is a valid time in the {1} zone.", _
                           localDate, _
                           TimeZoneInfo.Locale.StandardName)
    End If
   
    Dim localDate2 As DotNetLib.DateTime
    Set localDate2 = DateTime.FromBinary(binLocal)
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                        localDate, localDate2, localDate.Equals(localDate2))
End Sub

' The example displays the following output:
'    3/14/2010 2:30:00 AM is an invalid time in the Pacific Standard Time zone.
'    3/14/2010 2:30:00 AM = 3/14/2010 3:30:00 AM: False
'
' Or For eg. local timezone is AUS Eastern Standard Time zone
'    14/03/2010 2:30:00 AM is a valid time in the AUS Eastern Standard Time zone.
'    14/03/2010 2:30:00 AM = 14/03/2010 2:30:00 AM: True
