Attribute VB_Name = "DateTimeCreateFromTicksExample"
'@Folder "Examples.System.DateTime.Constructors"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 3, 2023
'@LastModified August 3, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.-ctor?view=netframework-4.8.1#system-datetime-ctor(system-int64)

Option Explicit

'@Description("This example demonstrates the DateTime(Int64) constructor.")
Public Sub DateTimeCreateFromTicks()
Attribute DateTimeCreateFromTicks.VB_Description = "This example demonstrates the DateTime(Int64) constructor."
    ' Instead of using the implicit, default "G" date and time format string, we
    ' use a custom format string that aligns the results and inserts leading zeroes.
    Dim pvtFormat As String
    pvtFormat = "MM/dd/yyyy hh:mm:ss tt"
    
    'Create a DateTime for the maximum date and time using ticks.
    Dim dt1 As IDateTime
    Set dt1 = DateTime.CreateFromTicks(DateTime.MaxValue.Ticks)
    
    'Create a DateTime for the minimum date and time using ticks.
    Dim dt2 As IDateTime
    Set dt2 = DateTime.CreateFromTicks(DateTime.MinValue.Ticks)
    
    'Create a custom DateTime for 7/28/1979 at 10:35:05 PM
    Dim pvtTicks As LongLong
    pvtTicks = DateTime.CreateFromDateTime(1979, 7, 28, 22, 35, 5).Ticks
    Dim dt3 As IDateTime
    Set dt3 = DateTime.CreateFromTicks(pvtTicks)
    
    Debug.Print "1) The maximum date and time is " & dt1.ToString2(pvtFormat)
    Debug.Print "2) The minimum date and time is " & dt2.ToString2(pvtFormat)
    Debug.Print "3) The custom  date and time is " & dt3.ToString2(pvtFormat)
    Debug.Print "The custom date and time is created from " & VBA.format$(pvtTicks, "#,##0") & " ticks."
End Sub

'/*
'This example produces the following results:
'
'1) The maximum date and time is 12/31/9999 11:59:59 PM
'2) The minimum date and time is 01/01/0001 12:00:00 AM
'3) The custom  date and time is 07/28/1979 10:35:05 PM
'
'The custom date and time is created from 624,376,461,050,000,000 ticks.
'
'*/
