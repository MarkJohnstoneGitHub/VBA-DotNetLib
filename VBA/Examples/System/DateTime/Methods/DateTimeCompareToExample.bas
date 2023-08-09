Attribute VB_Name = "DateTimeCompareToExample"
'@Folder("VBADotNetLib.Examples.DateTime.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 10, 2023
'@LastModified August 4, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.compareto?view=netframework-4.8.1#system-datetime-compareto(system-datetime)

Option Explicit

' The following example instantiates three DateTime objects, one that represents today's date,
' another that represents the date one year previously, and a third that represents the date
' one year in the future. It then calls the CompareTo(DateTime) method and displays the
' result of the comparison.

Private Enum DateComparisonResult
   Earlier = -1
   Later = 1
   TheSame = 0
End Enum

Public Sub DateTimeCompareTo()
   Dim thisDate As IDateTime
   Set thisDate = DateTime.Today
   
   ' Define two DateTime objects for today's date
   ' next year and last year
   Dim thisDateNextYear As IDateTime
   Dim thisDateLastYear As IDateTime

   ' Call AddYears instance method to add/substract 1 year
   Set thisDateNextYear = thisDate.AddYears(1)
   Set thisDateLastYear = thisDate.AddYears(-1)
   
   ' Compare dates
   Dim comparison As DateComparisonResult
   ' Compare today to last year
   comparison = thisDate.CompareTo(thisDateLastYear)
   Debug.Print "CompareTo method returns " & comparison & ": " & _
                thisDate.ToString2("d") & " is " & _
                DateComparisonResultToString(comparison) & " than " & thisDateLastYear.ToString2("d")
   
   ' Compare today to last year
   comparison = thisDate.CompareTo(thisDateNextYear)
   Debug.Print "CompareTo method returns " & comparison & ": " & _
                thisDate.ToString2("d") & " is " & _
                DateComparisonResultToString(comparison) & " than " & thisDateNextYear.ToString2("d")
End Sub

Private Function DateComparisonResultToString(ByVal value As DateComparisonResult) As String
   Select Case value
      Case Earlier: DateComparisonResultToString = "Earlier"
      Case Later: DateComparisonResultToString = "Later"
      Case TheSame: DateComparisonResultToString = "TheSame"
   End Select
End Function

Public Sub DateTimeCompareTo2()
   Dim theDay As IDateTime
   Set theDay = DateTime.CreateFromDate(DateTime.Today.Year, 7, 28)
   Dim compareValue As Long
   On Error GoTo ErrorHandler
      compareValue = theDay.CompareTo2(DateTime.Today)
   On Error GoTo 0
   If (compareValue < 0) Then
      Debug.Print theDay.ToString2("d") & " is in the past."
   ElseIf (compareValue = 0) Then
      Debug.Print theDay.ToString2("d") & " is today!"
   Else ' compareValue > 0
      Debug.Print theDay.ToString2("d") & " has not come yet."
   End If

CleanExit:
Exit Sub
    
ErrorHandler:
   Debug.Print "Value is not a DateTime"
   Debug.Print Err.Number, Err.Description
   Resume CleanExit
End Sub

' If run on October 20, 2006, the example produces the following output:
'    CompareTo method returns 1: 10/20/2006 is later than 10/20/2005
'    CompareTo method returns -1: 10/20/2006 is earlier than 10/20/2007

