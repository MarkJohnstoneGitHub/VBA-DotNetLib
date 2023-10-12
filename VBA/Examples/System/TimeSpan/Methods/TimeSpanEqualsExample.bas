Attribute VB_Name = "TimeSpanEqualsExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified August 1, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.equals?view=netframework-4.8.1#system-timespan-equals(system-timespan)

Option Explicit

'@Description("The following example demonstrates the TimeSpan.Equals method.")
Public Sub TimeSpanEquals()
Attribute TimeSpanEquals.VB_Description = "The following example demonstrates the TimeSpan.Equals method."
   ' Create some TimeSpan objects.
   Dim one As ITimeSpan
   Set one = TimeSpan.Create3(0, 0, 10, -20, -30)
   
   Dim two As ITimeSpan
   Set two = TimeSpan.Create3(0, -10, 20, -30, 40)
   
   Dim three As ITimeSpan
   Set three = one
   
   ' Compare the TimeSpan objects and display the results.
   Dim Result As Boolean
   Result = one.Equals(two)
   
   Debug.Print "The result of comparing TimeSpan object one and two is: " & Result & "."
   
   Result = one.Equals(three)
   Debug.Print "The result of comparing TimeSpan object one and three is: " & Result & "."
End Sub

' This code example displays the following:
'
' The result of comparing TimeSpan object one and two is: False.
' The result of comparing TimeSpan object one and three is: True.

   
