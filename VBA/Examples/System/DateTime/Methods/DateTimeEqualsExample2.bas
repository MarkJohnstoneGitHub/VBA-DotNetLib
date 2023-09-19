Attribute VB_Name = "DateTimeEqualsExample2"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 4, 2023
'@LastModified September 10, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.equals?view=netframework-4.8.1#system-datetime-equals(system-datetime-system-datetime)
Option Explicit

'@Description("Returns a value indicating whether two DateTime instances have the same date and time value.")
Public Sub DateTimeEquals3()
Attribute DateTimeEquals3.VB_Description = "Returns a value indicating whether two DateTime instances have the same date and time value."
   Dim today1 As IDateTime
   Set today1 = DateTime.CreateFromTicks(DateTime.Today.Ticks)
   
   Dim today2 As IDateTime
   Set today2 = DateTime.CreateFromTicks(DateTime.Today.Ticks)

   Dim tomorrow As IDateTime
   Set tomorrow = DateTime.CreateFromTicks(DateTime.Today.AddDays(1).Ticks)

   ' todayEqualsToday gets true.
   Dim todayEqualsToday As Boolean
   todayEqualsToday = DateTime.Equals(today1, today2)
   Debug.Print todayEqualsToday
   
   ' todayEqualsTomorrow gets false.
   Dim todayEqualsTomorrow As Boolean
   todayEqualsTomorrow = DateTime.Equals(today1, tomorrow)
   Debug.Print todayEqualsTomorrow
End Sub

' Output:
'    True
'    False
