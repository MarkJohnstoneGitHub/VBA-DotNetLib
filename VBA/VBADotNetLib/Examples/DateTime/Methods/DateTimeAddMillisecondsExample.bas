Attribute VB_Name = "DateTimeAddMillisecondsExample"
'@Folder("VBADotNetLib.Examples.DateTime.Methods")

'@DotNetReference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.addmilliseconds?view=netframework-4.8.1#examples

Option Explicit


'@Description("The following example uses the AddMilliseconds method to add one millisecond and 1.5 milliseconds to a DateTime value.")
' It then displays each new value and displays the difference between it and the original value.
' The difference is displayed both as a time span and as a number of ticks. The example makes it
' clear that one millisecond equals 10,000 ticks. It also shows that fractional milliseconds are
' rounded before performing the addition; the DateTime value that results from adding 1.5 milliseconds
' to the original date is 2 milliseconds greater than the original date.
Public Sub DateTimeAddMilliseconds()
   Dim dateFormat As String
   dateFormat = "MM/dd/yyyy hh:mm:ss.fffffff"
   Dim date1 As DateTime
   Set date1 = DateTime.CreateFromDateTime(2010, 9, 8, 16, 0, 0)
   Debug.Print "Original date: " & date1.ToString2(dateFormat) & " " & date1.Ticks & " ticks" & vbNewLine
   
   Dim date2 As DateTime
   Set date2 = date1.AddMilliseconds(1)
   Debug.Print "Second date:   " & date1.ToString2(dateFormat) & " " & date1.Ticks & " ticks"
   Debug.Print "Difference between dates: " & DateTime.Subtraction(date2, date1).ToString & " (" & date2.Ticks - date1.Ticks & " ticks)" & vbNewLine
   
   Dim date3 As DateTime
   Set date3 = date1.AddMilliseconds(1.5)
   Debug.Print "Third date:    " & date3.ToString2(dateFormat) & " " & date3.Ticks & " ticks"

   Debug.Print "Difference between dates: " & DateTime.Subtraction(date3, date1).ToString & " (" & date3.Ticks - date1.Ticks & " ticks)"
   
' The example displays the following output:
'    Original date: 09/08/2010 04:00:00.0000000 (634,195,584,000,000,000 ticks)
'
'    Second date:   09/08/2010 04:00:00.0010000 (634,195,584,000,010,000 ticks)
'    Difference between dates: 00:00:00.0010000 (10,000 ticks)
'
'    Third date:    09/08/2010 04:00:00.0020000 (634,195,584,000,020,000 ticks)
'    Difference between dates: 00:00:00.0020000 (20,000 ticks)

End Sub

