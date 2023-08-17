Attribute VB_Name = "DateTimeCompareTo2Example"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 5, 2023
'@LastModified August 5, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.compareto?view=netframework-4.8.1#system-datetime-compareto(system-object)

Option Explicit

'@Description("The following example demonstrates the CompareTo method.")
Public Sub DateTimeCompareTo2()
Attribute DateTimeCompareTo2.VB_Description = "The following example demonstrates the CompareTo method."
    Dim theDay As DotNetLib.DateTime
    Set theDay = DateTime.CreateFromDate(DateTime.Today.Year, 7, 28)
    Dim compareValue As Boolean
    
    On Error Resume Next
    compareValue = theDay.CompareTo2(DateTime.Today)
    If Catch(ArgumentException) Then
        Debug.Print "Value is not a DateTime"
        On Error GoTo 0 'Stop code and display error
        Exit Sub
    End If
    If (compareValue < 0) Then
        Debug.Print theDay.ToString2("d") & " is in the past."
    ElseIf (compareValue = 0) Then
        Debug.Print theDay.ToString2("d") & " is today!"
    Else
        Debug.Print theDay.ToString2("d") & " has not come yet."
    End If
End Sub

' Output:
'   7/28/2023 is in the past.
