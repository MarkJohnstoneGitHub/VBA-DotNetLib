Attribute VB_Name = "DateTimeEqualsExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 10, 2023
'@LastModified August 4, 2023

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
   Dim Result As Boolean
   Result = one.Equals(two)
   
   Debug.Print "The result of comparing DateTime object one and two is: " & Result & "."
   
   Result = one.Equals(three)
   Debug.Print "The result of comparing DateTime object one and three is: " & Result & "."
End Sub

' This code example displays the following:
'
' The result of comparing DateTime object one and two is: False.
' The result of comparing DateTime object one and three is: True.
