Attribute VB_Name = "DateTimeDayOfWeekExample"
'@Folder("Examples.DateTime")
Option Explicit

' https://learn.microsoft.com/en-us/dotnet/api/system.datetime.dayofweek?view=netframework-4.8.1#examples
'@Description("The following example demonstrates the DayOfWeek property and the DayOfWeek enumeration."
Public Sub DateTimeDayOfWeek()
    ' Assume the current culture is en-US.
    ' Create a DateTime for the first of May, 2003.
    Dim dt As DateTime
    Set dt = DateTime.CreateFromDate(2003, 5, 1)
    Debug.Print "Is Thursday the day of the week for " & dt.ToString & "?: " & IIf(dt.DayOfWeek = DayOfWeek.DayOfWeek_Thursday, True, False)
    Debug.Print "The day of the week for " & dt.ToString2("d") & " is " & DayOfWeekHelper.ToString(dt.DayOfWeek)

'This example produces the following results:
'
'Is Thursday the day of the week for 5/1/2003?: True
'The day of the week for 5/1/2003 is Thursday.

End Sub


