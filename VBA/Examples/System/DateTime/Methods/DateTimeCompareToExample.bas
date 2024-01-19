Attribute VB_Name = "DateTimeCompareToExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 10, 2023
'@LastModified January 6, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.compareto?view=netframework-4.8.1#system-datetime-compareto(system-datetime)

Option Explicit

''
' The following example instantiates three DateTime objects, one that represents today's date,
' another that represents the date one year previously, and a third that represents the date
' one year in the future. It then calls the CompareTo(DateTime) method and displays the
' result of the comparison.
''

Private Enum DateComparisonResult
   Earlier = -1
   Later = 1
   TheSame = 0
End Enum

Public Sub DateTimeCompareTo()
    Dim thisDate As DotNetLib.DateTime
    Set thisDate = DateTime.Today
   
    ' Define two DateTime objects for today's date
    ' next year and last year
    Dim thisDateNextYear As DotNetLib.DateTime
    Dim thisDateLastYear As DotNetLib.DateTime

    ' Call AddYears instance method to add/substract 1 year
    Set thisDateNextYear = thisDate.AddYears(1)
    Set thisDateLastYear = thisDate.AddYears(-1)
   
    ' Compare dates
    Dim pvtComparison As DateComparisonResult
    ' Compare today to last year
    pvtComparison = thisDate.CompareTo(thisDateLastYear)
    Debug.Print VBString.Format("CompareTo method returns {0}: {1:d} is {2} than {3:d}", _
                        pvtComparison, thisDate, LCase$(DateComparisonResultToString(pvtComparison)), _
                        thisDateLastYear)
   
    ' Compare today to next year
    pvtComparison = thisDate.CompareTo(thisDateNextYear)
    Debug.Print VBString.Format("CompareTo method returns {0}: {1:d} is {2} than {3:d}", _
                        pvtComparison, thisDate, LCase$(DateComparisonResultToString(pvtComparison)), _
                        thisDateNextYear)
End Sub

Private Function DateComparisonResultToString(ByVal value As DateComparisonResult) As String
   Select Case value
      Case Earlier: DateComparisonResultToString = "Earlier"
      Case Later: DateComparisonResultToString = "Later"
      Case TheSame: DateComparisonResultToString = "TheSame"
   End Select
End Function

' If run on October 20, 2006, the example produces the following output:
'    CompareTo method returns 1: 10/20/2006 is later than 10/20/2005
'    CompareTo method returns -1: 10/20/2006 is earlier than 10/20/2007


