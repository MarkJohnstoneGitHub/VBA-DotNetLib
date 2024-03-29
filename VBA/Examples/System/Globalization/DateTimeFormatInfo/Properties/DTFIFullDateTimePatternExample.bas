Attribute VB_Name = "DTFIFullDateTimePatternExample"
'@Folder "Examples.System.Globalization.DateTimeFormatInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 1, 2023
'@LastModified September 2, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.fulldatetimepattern?view=netframework-4.8.1#examples

Option Explicit

' The following example displays the value of FullDateTimePattern for a few cultures.
Public Sub DateTimeFormatInfoFullDateTimePattern()
    ' Displays the values of the pattern properties.
    Debug.Print " CULTURE    PROPERTY VALUE"
    PrintPattern "en-US"
    PrintPattern "ja-JP"
    PrintPattern "fr-FR"
End Sub

Private Sub PrintPattern(ByVal myCulture As String)
    Dim myDtfi As DotNetLib.DateTimeFormatInfo
    Set myDtfi = CultureInfo.CreateFromName(myCulture, False).DateTimeFormat
    Debug.Print "  "; myCulture; "     "; myDtfi.FullDateTimePattern
    
    #If Not Mac Then
        Dim messageBoxText As String
        messageBoxText = "  " & myCulture & "     " & myDtfi.FullDateTimePattern
        WinAPIUser32.MessageBoxW 0, StrPtr(messageBoxText), StrPtr("Culture FullDateTimePattern"), 0
    #End If
End Sub

'/*
'This code produces the following output.  The question marks take the place of native script characters.
'
' CULTURE    PROPERTY VALUE
'  en-US     dddd, MMMM dd, yyyy h:mm:ss tt
'  ja-JP     yyyy'?'M'?'d'?' H:mm:ss
'  fr-FR     dddd d MMMM yyyy HH:mm:ss
'
'*/
