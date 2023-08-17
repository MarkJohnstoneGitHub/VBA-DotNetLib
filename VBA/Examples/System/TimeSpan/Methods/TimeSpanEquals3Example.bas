Attribute VB_Name = "TimeSpanEquals3Example"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified August 14, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.equals?view=netframework-4.8.1#system-timespan-equals(system-timespan-system-timespan)

Option Explicit

'@Description("The following example demonstrates the TimeSpan.Equals( TimeSpan, TimeSpan ) method.")
Public Sub TimeSpanEquals3()
Attribute TimeSpanEquals3.VB_Description = "The following example demonstrates the TimeSpan.Equals( TimeSpan, TimeSpan ) method."
   ' Create some TimeSpan objects.
   Dim one As ITimeSpan
   Set one = TimeSpan.Create3(0, 0, 10, -20, -30)
   
   Dim two As ITimeSpan
   Set two = TimeSpan.Create3(0, -10, 20, -30, 40)
   
   Dim three As ITimeSpan
   Set three = one
   
   ' Compare the TimeSpan objects and display the results.
   Dim result As Boolean
   result = TimeSpan.Equals(one, two)
   
   Debug.Print "The result of comparing TimeSpan object one and two is: " & result & "."
   
   result = TimeSpan.Equals(one, three)
   Debug.Print "The result of comparing TimeSpan object one and three is: " & result & "."
End Sub

' This code example displays the following:
'
' The result of comparing TimeSpan object one and two is: False.
' The result of comparing TimeSpan object one and three is: True.
