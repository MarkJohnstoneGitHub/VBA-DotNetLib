Attribute VB_Name = "DateTimeAddDaysExample"
'@Folder("Examples.DateTime.Methods")

Option Explicit

'@Description("The following example uses the AddDays method to determine the day of the week 36 days after the current date.")
Public Sub DateTimeAddDays()
   Dim pvtToday As DateTime
   Set pvtToday = DateTime.Now
   Dim answer As DateTime
   Set answer = pvtToday.AddDays(36)
   Debug.Print "Today: " & pvtToday.ToString2("dddd")
   Debug.Print "36 days from today: " & answer.ToString2("dddd")
   
' The example displays output like the following:
'       Today: Wednesday
'       36 days from today: Thursday

End Sub
