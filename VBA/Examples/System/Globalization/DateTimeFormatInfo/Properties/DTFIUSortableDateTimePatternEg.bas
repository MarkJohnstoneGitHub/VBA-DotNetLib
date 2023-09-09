Attribute VB_Name = "DTFIUSortableDateTimePatternEg"
'@Folder("Examples.System.Globalization.DateTimeFormatInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 4, 2023
'@LastModified September 4, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.universalsortabledatetimepattern?view=netframework-4.8.1#examples

Option Explicit

' The following example displays the value of UniversalSortableDateTimePattern for a few cultures.
Public Sub DateTimeFormatInfoUniversalSortableDateTimePattern()
    ' Displays the values of the pattern properties.
    Debug.Print " CULTURE    PROPERTY VALUE"
    PrintPattern "en-US"
    PrintPattern "ja-JP"
    PrintPattern "fr-FR"
End Sub

Private Sub PrintPattern(ByVal myCulture As String)
    Dim myDtfi As DotNetLib.DateTimeFormatInfo
    Set myDtfi = CultureInfo.CreateFromName(myCulture, False).DateTimeFormat
    Debug.Print "  "; myCulture; "     "; myDtfi.UniversalSortableDateTimePattern
End Sub

'/*
'This code produces the following output.
'
' CULTURE    PROPERTY VALUE
'  en-US     yyyy'-'MM'-'dd HH':'mm':'ss'Z'
'  ja-JP     yyyy'-'MM'-'dd HH':'mm':'ss'Z'
'  fr-FR     yyyy'-'MM'-'dd HH':'mm':'ss'Z'
'
'*/
