Attribute VB_Name = "DateTimeMinValueExample"
'Rubberduck annotations
'@Folder "Examples.System.DateTime.Fields"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 9, 2023
'@LastModified August 3, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.minvalue?view=netframework-4.8.1

Option Explicit

' The value of this constant is equivalent to 00:00:00.0000000 UTC, January 1, 0001, in the Gregorian calendar.
' MinValue defines the date and time that is assigned to an uninitialized DateTime variable.
' The following example illustrates this.
Public Sub DateTimeMinValue()
    Dim date1 As IDateTime
    Set date1 = New DotNetLib.DateTime
    If (date1.Equals(DateTime.MinValue)) Then
        Debug.Print date1.ToString & " (Equals Date.MinValue)"
    End If
End Sub

' The example displays the following output:
'    1/1/0001 12:00:00 AM  (Equals Date.MinValue)
