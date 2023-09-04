Attribute VB_Name = "DTFIYearMonthPatternExample"
'@Folder("Examples.System.Globalization.DateTimeFormatInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 4, 2023
'@LastModified September 4, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.yearmonthpattern?view=netframework-4.8.1#examples

Option Explicit

' The following example displays the value of YearMonthPattern for a few cultures.
Public Sub DateTimeFormatInfoYearMonthPattern()
    ' Displays the values of the pattern properties.
    Debug.Print " CULTURE    PROPERTY VALUE"
    PrintPattern "en-US"
    PrintPattern "ja-JP"
    PrintPattern "fr-FR"
End Sub

Private Sub PrintPattern(ByVal myCulture As String)
    Dim myDTFI As DotNetLib.DateTimeFormatInfo
    Set myDTFI = CultureInfo.CreateFromName(myCulture, False).DateTimeFormat
    Debug.Print "  "; myCulture; "     "; myDTFI.YearMonthPattern
End Sub
