Attribute VB_Name = "DateTimeEquals2Example"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 4, 2023
'@LastModified August 4, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.equals?view=netframework-4.8.1#system-datetime-equals(system-object)

Option Explicit

'@Description("The following example uses the Equals(Object) method to determine whether the DateTime instances are equal.")
Public Sub DateTimeEquals2()
Attribute DateTimeEquals2.VB_Description = "The following example uses the Equals(Object) method to determine whether the DateTime instances are equal."
   ' Create some DateTime objects.
   Dim one As IDateTime
   Set one = DateTime.UtcNow
   
   Dim two As Object
   Set two = DateTime.Now
   
   Dim three As Object
   Set three = one
   
   ' Compare the DateTime objects and display the results.
   Dim Result As Boolean
   Result = one.Equals2(two)
   
   Debug.Print "The result of comparing DateTime object one and two is: " & Result & "."
   
   Result = one.Equals2(three)
   Debug.Print "The result of comparing DateTime object one and three is: " & Result & "."
End Sub

' This code example displays the following:
'
' The result of comparing DateTime object one and two is: False.
' The result of comparing DateTime object one and three is: True.
