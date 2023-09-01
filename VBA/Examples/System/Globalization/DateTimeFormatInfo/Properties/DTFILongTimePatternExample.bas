Attribute VB_Name = "DTFILongTimePatternExample"
'@Folder("Examples.System.Globalization.DateTimeFormatInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 1, 2023
'@LastModified September 1, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.longtimepattern?view=netframework-4.8.1#examples

Option Explicit

' The following example displays the value of LongTimePattern for a few cultures.
Public Sub DateTimeFormatInfoLongTimePattern()
    ' Displays the values of the pattern properties.
    Debug.Print " CULTURE    PROPERTY VALUE"
    PrintPattern "en-US"
    PrintPattern "ja-JP"
    PrintPattern "fr-FR"
End Sub

Private Sub PrintPattern(ByVal myCulture As String)
    Dim myDTFI As DotNetLib.DateTimeFormatInfo
    Set myDTFI = CultureInfo.Create4(myCulture, False).DateTimeFormat
    Debug.Print "  "; myCulture; "     "; myDTFI.LongTimePattern
End Sub

'/*
'This code produces the following output.
'
' CULTURE    PROPERTY VALUE
'  en-US     h:mm:ss tt
'  ja-JP     H:mm:ss
'  fr-FR     HH:mm:ss
'
'*/
