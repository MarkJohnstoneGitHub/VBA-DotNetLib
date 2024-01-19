Attribute VB_Name = "DateTimeAddSecondsExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 10, 2023
'@LastModified January 6, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.addseconds?view=netframework-4.8.1#examples

Option Explicit

' The following example uses the AddSeconds method to add 30 seconds and the number of seconds in one day to a DateTime value.
' It then displays each new value and displays the difference between it and the original value. The difference is displayed both as a time span and as a number of ticks.
Public Sub DateTimeAddSeconds()
    Dim dateFormat As String
    dateFormat = "MM/dd/yyyy hh:mm:ss"
    
    Dim date1 As DotNetLib.DateTime
    Set date1 = DateTime.CreateFromDateTime(2014, 9, 8, 16, 0, 0)
    Debug.Print VBString.Format(VBString.Unescape("Original date: {0} ({1:N0} ticks)\n"), _
                        date1.ToString2(dateFormat), date1.Ticks)
                
    Dim date2 As DotNetLib.DateTime
    Set date2 = date1.AddSeconds(30)
    
    Debug.Print VBString.Format("Second date:   {0} ({1:N0} ticks)", _
                        date2.ToString2(dateFormat), date2.Ticks)
    Debug.Print VBString.Format(VBString.Unescape("Difference between dates: {0} ({1:N0} ticks)\n"), _
                        DateTime.Subtraction(date2, date1), date2.Ticks - date1.Ticks)
    
    ' Add 1 day's worth of seconds (60 secs. * 60 mins * 24 hrs.
    Dim date3 As DotNetLib.DateTime
    Set date3 = date1.AddSeconds(CDbl(60) * 60 * 24) 'convert to double to avoid VBA overflow error
    Debug.Print VBString.Format("Third date:    {0} ({1:N0} ticks)", _
                        date3.ToString2(dateFormat), date3.Ticks)
    Debug.Print VBString.Format("Difference between dates: {0} ({1:N0} ticks)", _
                        DateTime.Subtraction(date3, date1), date3.Ticks - date1.Ticks)
End Sub

' The example displays the following output:
'    Original date: 09/08/2014 04:00:00 (635,457,888,000,000,000 ticks)
'
'    Second date:   09/08/2014 04:00:30 (635,457,888,300,000,000 ticks)
'    Difference between dates: 00:00:30 (300,000,000 ticks)
'
'    Third date:    09/09/2014 04:00:00 (635,458,752,000,000,000 ticks)
'    Difference between dates: 1.00:00:00 (864,000,000,000 ticks)


