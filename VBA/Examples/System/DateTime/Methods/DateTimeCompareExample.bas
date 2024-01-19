Attribute VB_Name = "DateTimeCompareExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 10, 2023
'@LastModified January 6, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.compare?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example demonstrates the Compare method.")
Public Sub DateTimeCompare()
Attribute DateTimeCompare.VB_Description = "The following example demonstrates the Compare method."
    Dim date1 As DotNetLib.DateTime
    Set date1 = DateTime.CreateFromDateTime(2009, 8, 1, 0, 0, 0)
    Dim date2 As DotNetLib.DateTime
    Set date2 = DateTime.CreateFromDateTime(2009, 8, 1, 12, 0, 0)
    
    Dim result As Long
    result = DateTime.Compare(date1, date2)
    Dim relationship As String
    
    If (result < 0) Then
        relationship = "is earlier than"
    ElseIf (result = 0) Then
        relationship = "is the same time as"
    Else
        relationship = "is later than"
    End If
   
    Debug.Print VBString.Format("{0} {1} {2}", date1, relationship, date2)
End Sub

' The example displays the following output for en-us culture:
'    8/1/2009 12:00:00 AM is earlier than 8/1/2009 12:00:00 PM
