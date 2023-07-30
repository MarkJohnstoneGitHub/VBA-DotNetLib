Attribute VB_Name = "DateTimeEqualsExample"
'@Folder("VBADotNetLib.Examples.DateTime.Methods")

'https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 10, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.equals?view=netframework-4.8.1#system-datetime-equals(system-datetime)

Option Explicit

'@Description("The following example demonstrates the Equals method.")
Public Sub DateTimeEquals()
Attribute DateTimeEquals.VB_Description = "The following example demonstrates the Equals method."

   ' Create some DateTime objects.
   Dim one As IDateTime
   Set one = DateTime.UtcNow
   
   Dim two As IDateTime
   Set two = DateTime.Now
   
   Dim three As IDateTime
   Set three = one
   
   ' Compare the DateTime objects and display the results.
   Dim result As Boolean
   result = one.Equals(two)
   
   Debug.Print "The result of comparing DateTime object one and two is: " & result & "."
   
   result = one.Equals(three)
   Debug.Print "The result of comparing DateTime object one and three is: " & result & "."
End Sub

' This code example displays the following:
'
' The result of comparing DateTime object one and two is: False.
' The result of comparing DateTime object one and three is: True.

Public Sub DateTimeEquals3()
   Dim today1 As IDateTime
   Set today1 = DateTime.CreateFromTicks(DateTime.Today.Ticks)
   
   Dim today2 As IDateTime
   Set today2 = DateTime.CreateFromTicks(DateTime.Today.Ticks)

   Dim tomorrow As IDateTime
   Set tomorrow = DateTime.CreateFromTicks(DateTime.Today.AddDays(1).Ticks)

   ' todayEqualsToday gets true.
   Dim todayEqualsToday As Boolean
   todayEqualsToday = DateTime.Equals3(today1, today2)
   
   ' todayEqualsTomorrow gets false.
   Dim todayEqualsTomorrow As Boolean
   todayEqualsTomorrow = DateTime.Equals3(today1, tomorrow)

End Sub
