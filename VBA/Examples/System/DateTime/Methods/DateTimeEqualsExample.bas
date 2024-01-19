Attribute VB_Name = "DateTimeEqualsExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 10, 2023
'@LastModified January 6, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.equals?view=netframework-4.8.1#system-datetime-equals(system-datetime)

Option Explicit

'@Description("The following example demonstrates the Equals method.")
Public Sub DateTimeEquals()
Attribute DateTimeEquals.VB_Description = "The following example demonstrates the Equals method."
    ' Create some DateTime objects.
    Dim one As DotNetLib.DateTime
    Set one = DateTime.UtcNow
    
    Dim two As DotNetLib.DateTime
    Set two = DateTime.Now
   
    Dim three As DotNetLib.DateTime
    Set three = one
   
    ' Compare the DateTime objects and display the results.
    Dim result As Boolean
    result = one.Equals(two)
   
    Debug.Print VBString.Format("The result of comparing DateTime object one and two is: {0}.", result)
   
    result = one.Equals(three)
    Debug.Print VBString.Format("The result of comparing DateTime object one and three is: {0}.", result)
End Sub

' This code example displays the following:
'
' The result of comparing DateTime object one and two is: False.
' The result of comparing DateTime object one and three is: True.
