Attribute VB_Name = "StringsFormatExample2"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 23, 2023
'@LastModified September 23, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.format?view=netframework-4.8.1#system-string-format(system-string-system-object)

Option Explicit

' The following example uses the Format(String, Object) method to embed an
' individual's age in the middle of a string.
Public Sub StringsFormatEg2()
    Dim birthDate As DotNetLib.DateTime
    Set birthDate = DateTime.CreateFromDate(1993, 7, 28)
    Dim dates() As DotNetLib.DateTime
    
    ObjectArray.CreateInitialize1D dates, _
                    DateTime.CreateFromDate(1993, 8, 16), _
                    DateTime.CreateFromDate(1994, 7, 28), _
                    DateTime.CreateFromDate(2000, 10, 16), _
                    DateTime.CreateFromDate(2003, 7, 27), _
                    DateTime.CreateFromDate(2007, 5, 27)
    
    Dim varDateValue As Variant
    For Each varDateValue In dates
        Dim dateValue As DotNetLib.DateTime
        Set dateValue = varDateValue

        Dim interval As DotNetLib.TimeSpan
        Set interval = DateTime.Subtraction(dateValue, birthDate)
        
        ' Get the approximate number of years, without accounting for leap years.
        Dim pvtYears As Long
        pvtYears = interval.TotalDays / 365
        
        ' See if adding the number of years exceeds dateValue.
        Dim Output As String
        If (DateTime.LessThanOrEqual(birthDate.AddYears(pvtYears), dateValue)) Then
            Output = VBString.Format("You are now {0} years old.", pvtYears)
            Debug.Print Output
        Else
            Output = VBString.Format("You are now {0} years old.", pvtYears - 1)
            Debug.Print Output
        End If
    Next
End Sub

' The example displays the following output:
'       You are now 0 years old.
'       You are now 1 years old.
'       You are now 7 years old.
'       You are now 9 years old.
'       You are now 13 years old.


