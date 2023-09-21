Attribute VB_Name = "DTOCreateFromDateTimeParts3Eg"
'@Folder("Examples.System.DateTimeOffset.Constructors")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 21, 2023
'@LastModified September 21, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.-ctor?view=netframework-4.8.1#system-datetimeoffset-ctor(system-int32-system-int32-system-int32-system-int32-system-int32-system-int32-system-int32-system-globalization-calendar-system-timespan)

Option Explicit

' The following example uses instances of both the HebrewCalendar class and the
' HijriCalendar class to instantiate a DateTimeOffset value. That date is then
' displayed to the console using the respective calendars and the Gregorian calendar.
Public Sub DateTimeOffsetCreateFromDateTimeParts3()
    Dim fmt As DotNetLib.CultureInfo
    Dim pvtYear As Long
    Dim cal As DotNetLib.ICalendar
    Dim dateInCal As DotNetLib.DateTimeOffset
    
    ' Instantiate DateTimeOffset with Hebrew calendar
    pvtYear = 5770
    Set cal = New DotNetLib.HebrewCalendar
    Set fmt = CultureInfo.CreateFromName("he-IL")
    Set fmt.DateTimeFormat.Calendar = cal
    Set dateInCal = DateTimeOffset.CreateFromDateTimeParts3(pvtYear, 7, 12, 15, 30, 0, 0, cal, TimeSpan.Create(2, 0, 0))

    ' Display the date in the Hebrew calendar
    Debug.Print "Date in Hebrew Calendar: "; dateInCal.ToString2("g", fmt)
    DisplayMessage "Date in Hebrew Calendar:", dateInCal.ToString2("g", fmt)
    
    ' Display the date in the Gregorian calendar
    Debug.Print "Date in Gregorian Calendar:  "; dateInCal.ToString2("g")
    DisplayMessage "Date in Gregorian Calendar: ", dateInCal.ToString2("g")
    
    ' Instantiate DateTimeOffset with Hijri calendar
    pvtYear = 1431
    Set cal = New DotNetLib.HijriCalendar
    Set fmt = CultureInfo.CreateFromName("ar-SA")
    Set fmt.DateTimeFormat.Calendar = cal
    Set dateInCal = DateTimeOffset.CreateFromDateTimeParts3(pvtYear, 7, 12, 15, 30, 0, 0, cal, TimeSpan.Create(2, 0, 0))
    
    ' Display the date in the Hijri calendar
    Debug.Print "Date in Hijri Calendar:  "; dateInCal.ToString2("g", fmt)
    DisplayMessage "Date in Hijri Calendar: ", dateInCal.ToString2("g", fmt)
    
    ' Display the date in the Gregorian calendar
    Debug.Print "Date in Gregorian Calendar:  "; dateInCal.ToString2("g")
    DisplayMessage "Date in Gregorian Calendar: ", dateInCal.ToString2("g")
    
End Sub


Private Sub DisplayMessage(ByVal title As String, ByVal messsage As String)
    #If Not Mac Then
        WinAPIUser32.MessageBoxW 0, StrPtr(messsage), StrPtr(title), 0
    #End If
End Sub
