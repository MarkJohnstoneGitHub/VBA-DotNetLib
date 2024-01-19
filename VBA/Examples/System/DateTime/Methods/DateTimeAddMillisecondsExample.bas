Attribute VB_Name = "DateTimeAddMillisecondsExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 10, 2023
'@LastModified January 6, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.addmilliseconds?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example uses the AddMilliseconds method to add one millisecond and 1.5 milliseconds to a DateTime value.")
' It then displays each new value and displays the difference between it and the original value.
' The difference is displayed both as a time span and as a number of ticks. The example makes it
' clear that one millisecond equals 10,000 ticks. It also shows that fractional milliseconds are
' rounded before performing the addition; the DateTime value that results from adding 1.5 milliseconds
' to the original date is 2 milliseconds greater than the original date.
Public Sub DateTimeAddMilliseconds()
Attribute DateTimeAddMilliseconds.VB_Description = "The following example uses the AddMilliseconds method to add one millisecond and 1.5 milliseconds to a DateTime value."
    Dim dateFormat As String
    dateFormat = "MM/dd/yyyy hh:mm:ss.fffffff"
    Dim date1 As DotNetLib.DateTime
    Set date1 = DateTime.CreateFromDateTime(2010, 9, 8, 16, 0, 0)
    Debug.Print VBString.Format(VBString.Unescape("Original date: {0} ({1:N0} ticks)\n"), _
                      date1.ToString2(dateFormat), date1.Ticks)
   
    Dim date2 As DotNetLib.DateTime
    Set date2 = date1.AddMilliseconds(1)
    Debug.Print VBString.Format("Second date:   {0} ({1:N0} ticks)", _
                  date2.ToString2(dateFormat), date2.Ticks)
    Debug.Print VBString.Format(VBString.Unescape("Difference between dates: {0} ({1:N0} ticks)\n"), _
                  DateTime.Subtraction(date2, date1), date2.Ticks - date1.Ticks)
   
    Dim date3 As DotNetLib.DateTime
    Set date3 = date1.AddMilliseconds(1.5)
    Debug.Print VBString.Format("Third date:    {0} ({1:N0} ticks)", _
                      date3.ToString2(dateFormat), date3.Ticks)
    Debug.Print VBString.Format("Difference between dates: {0} ({1:N0} ticks)", _
                      DateTime.Subtraction(date3, date1), date3.Ticks - date1.Ticks)
End Sub

' The example displays the following output:
'    Original date: 09/08/2010 04:00:00.0000000 (634,195,584,000,000,000 ticks)
'
'    Second date:   09/08/2010 04:00:00.0010000 (634,195,584,000,010,000 ticks)
'    Difference between dates: 00:00:00.0010000 (10,000 ticks)
'
'    Third date:    09/08/2010 04:00:00.0020000 (634,195,584,000,020,000 ticks)
'    Difference between dates: 00:00:00.0020000 (20,000 ticks)
