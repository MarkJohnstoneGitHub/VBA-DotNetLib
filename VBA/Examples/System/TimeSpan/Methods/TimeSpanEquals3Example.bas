Attribute VB_Name = "TimeSpanEquals3Example"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified January 17, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.equals?view=netframework-4.8.1#system-timespan-equals(system-timespan-system-timespan)

Option Explicit

''
' The following example demonstrates the
' TimeSpan.Equals( TimeSpan, TimeSpan ) method.
''
Public Sub TimeSpanEquals3()
   ' Create some TimeSpan objects.
   Dim one As DotNetLib.TimeSpan
   Set one = TimeSpan.Create3(0, 0, 10, -20, -30)
   
   Dim two As DotNetLib.TimeSpan
   Set two = TimeSpan.Create3(0, -10, 20, -30, 40)
   
   Dim three As DotNetLib.TimeSpan
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
