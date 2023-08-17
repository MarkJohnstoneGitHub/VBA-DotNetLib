Attribute VB_Name = "DateTimeOffsetTicksExample"
'@Folder "Examples.System.DateTimeOffset.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 19, 2023
'@LastModified August 4,2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.ticks?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example initializes a DateTimeOffset object by approximating the number of ticks in the date July 1, 2008 1:23:07. It then displays the date and the number of ticks represented by that date.")
Public Sub DateTimeOffsetTicks()
Attribute DateTimeOffsetTicks.VB_Description = "The following example initializes a DateTimeOffset object by approximating the number of ticks in the date July 1, 2008 1:23:07. It then displays the date and the number of ticks represented by that date."
    ' Attempt to initialize date to number of ticks
    ' in July 1, 2008 1:23:07.
    
    ' There are 10,000,000 100-nanosecond intervals in a second
    Const NSPerSecond As LongLong = 10000000
    Dim pvtTicks As LongLong
    pvtTicks = 7 * NSPerSecond                                                 ' Ticks in a 7 seconds
    pvtTicks = pvtTicks + (23 * 60 * NSPerSecond)                              ' Ticks in 23 minutes
    pvtTicks = pvtTicks + (1 * 60 * 60 * NSPerSecond)                          ' Ticks in 1 hour
    pvtTicks = pvtTicks + (CLngLng(60) * 60 * 24 * NSPerSecond)                ' Ticks in 1 day
    pvtTicks = pvtTicks + (CLngLng(181) * 60 * 60 * 24 * NSPerSecond)          ' Ticks in 6 months
    pvtTicks = pvtTicks + (CLngLng(2007) * 60 * 60 * 24 * 365 * NSPerSecond)   ' Ticks in 2007 years
    pvtTicks = pvtTicks + (CLngLng(486) * 60 * 60 * 24 * NSPerSecond)          ' Adjustment for leap years
    
    Dim dto As IDateTimeOffset
    Set dto = DateTimeOffset.CreateFromTicks(pvtTicks, DateTimeOffset.Now.Offset)
    Debug.Print "There are " & VBA.format$(dto.Ticks, "#,###") & " ticks in " & dto.ToString() & "."
End Sub

' The example displays the following output:
'       There are 633,504,721,870,000,000 ticks in 7/1/2008 1:23:07 AM -08:00.
