Attribute VB_Name = "DTFILongDatePatternExample"
'@Folder("Examples.System.Globalization.DateTimeFormatInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 1, 2023
'@LastModified September 1, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.longdatepattern?view=netframework-4.8.1#examples

Option Explicit

' The following example displays the value of the LongDatePattern property for a few cultures.
Public Sub DateTimeFormatInfoLongDatePattern()
    ' Displays the values of the pattern properties.
    Debug.Print " CULTURE    PROPERTY VALUE"
    PrintPattern "en-US"
    PrintPattern "ja-JP"
    PrintPattern "fr-FR"
End Sub

Private Sub PrintPattern(ByVal myCulture As String)
    Dim myDTFI As DotNetLib.DateTimeFormatInfo
    Set myDTFI = CultureInfo.Create4(myCulture, False).DateTimeFormat
    Debug.Print "  "; myCulture; "     "; myDTFI.LongDatePattern
    
    #If Not Mac Then
        Dim messageBoxText As String
        messageBoxText = "  " & myCulture & "     " & myDTFI.LongDatePattern
        WinAPIUser32.MessageBoxW 0, StrPtr(messageBoxText), StrPtr("Culture LongDatePattern"), 0
    #End If
End Sub

'/*
'This code produces the following output:
'
' CULTURE    PROPERTY VALUE
'  en-US     dddd, MMMM d, yyyy
'  ja-JP     yyyy'年'M'月'd'日'
'  fr-FR     dddd d MMMM yyyy
'
'*/
