Attribute VB_Name = "DTFISortableDateTimePatternEg"
'@Folder "Examples.System.Globalization.DateTimeFormatInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 4, 2023
'@LastModified September 4, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.sortabledatetimepattern?view=netframework-4.8.1#examples

Option Explicit

' The following example displays the value of SortableDateTimePattern for a few cultures.
Public Sub DateTimeFormatInfoSortableDateTimePattern()
    ' Displays the values of the pattern properties.
    Debug.Print " CULTURE    PROPERTY VALUE"
    PrintPattern "en-US"
    PrintPattern "ja-JP"
    PrintPattern "fr-FR"
End Sub

Private Sub PrintPattern(ByVal myCulture As String)
    Dim myDtfi As DotNetLib.DateTimeFormatInfo
    Set myDtfi = CultureInfo.CreateFromName(myCulture, False).DateTimeFormat
    Debug.Print "  "; myCulture; "     "; myDtfi.SortableDateTimePattern
End Sub

'/*
'This code produces the following output.
'
' CULTURE    PROPERTY VALUE
'  en-US     yyyy'-'MM'-'dd'T'HH':'mm':'ss
'  ja-JP     yyyy'-'MM'-'dd'T'HH':'mm':'ss
'  fr-FR     yyyy'-'MM'-'dd'T'HH':'mm':'ss
'
'*/
