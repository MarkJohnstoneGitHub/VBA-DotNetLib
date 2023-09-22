Attribute VB_Name = "DateTimeAddSecondsExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 10, 2023
'@LastModified August 4, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.addseconds?view=netframework-4.8.1#examples

Option Explicit

' The following example uses the AddSeconds method to add 30 seconds and the number of seconds in one day to a DateTime value.
' It then displays each new value and displays the difference between it and the original value. The difference is displayed both as a time span and as a number of ticks.
Public Sub DateTimeAddSeconds()
    Dim dateFormat As String
    dateFormat = "MM/dd/yyyy hh:mm:ss"
    
    Dim date1 As IDateTime
    Set date1 = DateTime.CreateFromDateTime(2014, 9, 8, 16, 0, 0)
    Debug.Print "Original date: " & date1.ToString2(dateFormat) & _
                " (" & VBA.Format$(date1.Ticks, "#,##0") & " ticks)" & vbNewLine
   
    Dim date2 As IDateTime
    Set date2 = date1.AddSeconds(30)
    Debug.Print "Second date:   " & date2.ToString2(dateFormat) & _
                " (" & VBA.Format$(date2.Ticks, "#,##0") & " ticks)"
    Debug.Print "Difference between dates: " & DateTime.Subtraction(date2, date1).ToString & _
                 " (" & VBA.Format$(date2.Ticks - date1.Ticks, "#,##0") & " ticks)"
    
    ' Add 1 day's worth of seconds (60 secs. * 60 mins * 24 hrs.
    Dim date3 As IDateTime
    Set date3 = date1.AddSeconds(CDbl(60) * 60 * 24) 'convert to double to avoid VBA overflow error
    Debug.Print "Third date:    " & date3.ToString2(dateFormat) & _
                " (" & VBA.Format$(date3.Ticks, "#,##0") & " ticks)"
    Debug.Print "Difference between dates: " & DateTime.Subtraction(date3, date1).ToString & _
                " (" & VBA.Format$(date3.Ticks - date1.Ticks, "#,##0") & " ticks)"
End Sub

' The example displays the following output:
'    Original date: 09/08/2014 04:00:00 (635,457,888,000,000,000 ticks)
'
'    Second date:   09/08/2014 04:00:30 (635,457,888,300,000,000 ticks)
'    Difference between dates: 00:00:30 (300,000,000 ticks)
'
'    Third date:    09/09/2014 04:00:00 (635,458,752,000,000,000 ticks)
'    Difference between dates: 1.00:00:00 (864,000,000,000 ticks)


