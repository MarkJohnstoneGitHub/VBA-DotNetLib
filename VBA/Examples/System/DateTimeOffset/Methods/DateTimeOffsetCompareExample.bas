Attribute VB_Name = "DateTimeOffsetCompareExample"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 20, 2023
'@LastModified July 31, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.compare?view=netframework-4.8.1#examples

Option Explicit

Private Enum TimeComparison
   Earlier = -1
   Same = 0
   Later = 1
End Enum

'@Description("The following example illustrates calls to the Compare method to compare DateTimeOffset objects.")
Public Sub DateTimeOffsetCompare()
Attribute DateTimeOffsetCompare.VB_Description = "The following example illustrates calls to the Compare method to compare DateTimeOffset objects."
   Dim firstTime As IDateTimeOffset
   Set firstTime = DateTimeOffset.CreateFromDateTimeParts(2007, 9, 1, 6, 45, 0, TimeSpan.Create(-7, 0, 0))
   Dim secondTime As IDateTimeOffset
   Set secondTime = firstTime
   Debug.Print "Comparing " & _
            firstTime.ToString() & _
            " and " & _
            secondTime.ToString() & _
            ": " & _
            TimeComparisionToString(DateTimeOffset.Compare(firstTime, secondTime))
   
   Set secondTime = DateTimeOffset.CreateFromDateTimeParts(2007, 9, 1, 6, 45, 0, TimeSpan.Create(-6, 0, 0))
   Debug.Print "Comparing " & _
            firstTime.ToString() & _
            " and " & _
            secondTime.ToString() & _
            ": " & _
            TimeComparisionToString(DateTimeOffset.Compare(firstTime, secondTime))
            
   Set secondTime = DateTimeOffset.CreateFromDateTimeParts(2007, 9, 1, 8, 45, 0, TimeSpan.Create(-5, 0, 0))
   Debug.Print "Comparing " & _
            firstTime.ToString() & _
            " and " & _
            secondTime.ToString() & _
            ": " & _
            TimeComparisionToString(DateTimeOffset.Compare(firstTime, secondTime))
End Sub

Private Function TimeComparisionToString(ByVal comparison As Long) As String
   Select Case comparison
      Case Earlier: TimeComparisionToString = "Earlier"
      Case Same: TimeComparisionToString = "Same"
      Case Later: TimeComparisionToString = "Later"
   End Select
End Function

' The example displays the following output:
'       Comparing 9/1/2007 6:45:00 AM -07:00 and 9/1/2007 6:45:00 AM -07:00: Same
'       Comparing 9/1/2007 6:45:00 AM -07:00 and 9/1/2007 6:45:00 AM -06:00: Later
'       Comparing 9/1/2007 6:45:00 AM -07:00 and 9/1/2007 8:45:00 AM -05:00: Same
