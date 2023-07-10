Attribute VB_Name = "DateTimeCompareToExample"
'@Folder("VBADotNetLib.Examples.DateTime.Methods")

'https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 10, 2023
'@LastModified July 10, 2023

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
   Dim thisDate As DateTime
   Set thisDate = DateTime.Today
   
   ' Define two DateTime objects for today's date
   ' next year and last year
   Dim thisDateNextYear As DateTime
   Dim thisDateLastYear As DateTime

   ' Call AddYears instance method to add/substract 1 year
   Set thisDateNextYear = thisDate.AddYears(1)
   Set thisDateLastYear = thisDate.AddYears(-1)
   
   
   ' Compare dates
   Dim comparison As DateComparisonResult
   ' Compare today to last year
   comparison = thisDate.CompareTo(thisDateLastYear)
   Debug.Print "CompareTo method returns " & comparison & ": " & _
      thisDate.ToString2("d") & " is " & DateComparisonResultToString(comparison) & " than " & thisDateLastYear.ToString2("d")
   
   ' Compare today to last year
   comparison = thisDate.CompareTo(thisDateNextYear)
   Debug.Print "CompareTo method returns " & comparison & ": " & _
      thisDate.ToString2("d") & " is " & DateComparisonResultToString(comparison) & " than " & thisDateNextYear.ToString2("d")
      
' If run on October 20, 2006, the example produces the following output:
'    CompareTo method returns 1: 10/20/2006 is later than 10/20/2005
'    CompareTo method returns -1: 10/20/2006 is earlier than 10/20/2007
   
End Sub

Private Function DateComparisonResultToString(ByVal value As DateComparisonResult) As String
   Select Case value
      Case Earlier: DateComparisonResultToString = "Earlier"
      Case Later: DateComparisonResultToString = "Later"
      Case TheSame: DateComparisonResultToString = "TheSame"
   End Select
End Function
