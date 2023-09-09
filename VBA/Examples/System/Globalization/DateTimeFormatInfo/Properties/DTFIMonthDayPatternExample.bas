Attribute VB_Name = "DTFIMonthDayPatternExample"
'@Folder("Examples.System.Globalization.DateTimeFormatInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 1, 2023
'@LastModified September 2, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.monthdaypattern?view=netframework-4.8.1#examples

Option Explicit

' The following example displays the value of MonthDayPattern for a few cultures.
Public Sub DateTimeFormatInfoMonthDayPattern()
    ' Displays the values of the pattern properties.
    Debug.Print " CULTURE    PROPERTY VALUE"
    PrintPattern "en-US"
    PrintPattern "ja-JP"
    PrintPattern "fr-FR"
End Sub

Private Sub PrintPattern(ByVal myCulture As String)
    Dim myDtfi As DotNetLib.DateTimeFormatInfo
    Set myDtfi = CultureInfo.CreateFromName(myCulture, False).DateTimeFormat
    Debug.Print "  "; myCulture; "     "; myDtfi.MonthDayPattern
    
    #If Not Mac Then
        Dim messageBoxText As String
        messageBoxText = "  " & myCulture & "     " & myDtfi.MonthDayPattern
        WinAPIUser32.MessageBoxW 0, StrPtr(messageBoxText), StrPtr("Culture MonthDayPattern"), 0
    #End If
End Sub

'/*
'This code produces the following output.  The question marks take the place of native script characters.
'
' CULTURE    PROPERTY VALUE
'  en-US     MMMM dd
'  ja-JP     M'?'d'?'
'  fr-FR     d MMMM
'
'*/
