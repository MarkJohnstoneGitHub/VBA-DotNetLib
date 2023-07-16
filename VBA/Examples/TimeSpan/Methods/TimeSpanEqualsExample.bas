Attribute VB_Name = "TimeSpanEqualsExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 16, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.equals?view=netframework-4.8.1#system-timespan-equals(system-timespan)

Option Explicit

'@Description("The following example demonstrates the TimeSpan.Equals method.")
Public Sub TimeSpanEquals()
Attribute TimeSpanEquals.VB_Description = "The following example demonstrates the TimeSpan.Equals method."
   ' Create some TimeSpan objects.
   Dim one As TimeSpan
   Set one = TimeSpan.Create3(0, 0, 10, -20, -30)
   
   Dim two As TimeSpan
   Set two = TimeSpan.Create3(0, -10, 20, -30, 40)
   
   Dim three As TimeSpan
   Set three = one
   
   ' Compare the TimeSpan objects and display the results.
   Dim result As Boolean
   result = one.Equals(two)
   
   Debug.Print "The result of comparing TimeSpan object one and two is: " & result & "."
   
   result = one.Equals(three)
   Debug.Print "The result of comparing TimeSpan object one and three is: " & result & "."
   
' This code example displays the following:
'
' The result of comparing TimeSpan object one and two is: False.
' The result of comparing TimeSpan object one and three is: True.
End Sub


'@Description("The following example demonstrates the TimeSpan.Equals( TimeSpan, TimeSpan ) method.")
'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.equals?view=netframework-4.8.1#system-timespan-equals(system-timespan-system-timespan)
Public Sub TimeSpanEquals3()
Attribute TimeSpanEquals3.VB_Description = "The following example demonstrates the TimeSpan.Equals( TimeSpan, TimeSpan ) method."
   ' Create some TimeSpan objects.
   Dim one As TimeSpan
   Set one = TimeSpan.Create3(0, 0, 10, -20, -30)
   
   Dim two As TimeSpan
   Set two = TimeSpan.Create3(0, -10, 20, -30, 40)
   
   Dim three As TimeSpan
   Set three = one
   
   ' Compare the TimeSpan objects and display the results.
   Dim result As Boolean
   result = TimeSpan.Equals3(one, two)
   
   Debug.Print "The result of comparing TimeSpan object one and two is: " & result & "."
   
   result = TimeSpan.Equals3(one, three)
   Debug.Print "The result of comparing TimeSpan object one and three is: " & result & "."
   
' This code example displays the following:
'
' The result of comparing TimeSpan object one and two is: False.
' The result of comparing TimeSpan object one and three is: True.
End Sub

   
