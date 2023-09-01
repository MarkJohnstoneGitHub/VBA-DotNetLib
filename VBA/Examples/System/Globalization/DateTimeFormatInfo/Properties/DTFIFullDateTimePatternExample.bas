Attribute VB_Name = "DTFIFullDateTimePatternExample"
'@Folder("Examples.System.Globalization.DateTimeFormatInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 1, 2023
'@LastModified September 1, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.fulldatetimepattern?view=netframework-4.8.1#examples

Option Explicit

' http://blog.nkadesign.com/2013/10/01/vba-unicode-strings-and-the-windows-api/
#If Not Mac And VBA7 Then
    Private Declare PtrSafe Function MessageBoxU Lib "user32" Alias "MessageBoxW" _
        (ByVal hwnd As LongPtr, _
         ByVal lpText As LongPtr, _
         ByVal lpCaption As LongPtr, _
         ByVal wType As Long) As Long
#ElseIf Not Mac Then
    Private Declare Function MessageBoxU Lib "user32" Alias "MessageBoxW" _
        (ByVal hwnd As Long, _
         ByVal lpText As Long, _
         ByVal lpCaption As Long, _
         ByVal wType As Long) As Long
#End If

' The following example displays the value of FullDateTimePattern for a few cultures.
Public Sub DateTimeFormatInfoFullDateTimePattern()
    ' Displays the values of the pattern properties.
    Debug.Print " CULTURE    PROPERTY VALUE"
    PrintPattern "en-US"
    PrintPattern "ja-JP"
    PrintPattern "fr-FR"
End Sub

Private Sub PrintPattern(ByVal myCulture As String)
    Dim myDTFI As DotNetLib.DateTimeFormatInfo
    Set myDTFI = CultureInfo.Create4(myCulture, False).DateTimeFormat
    Debug.Print "  "; myCulture; "     "; myDTFI.FullDateTimePattern
    
    #If Not Mac Then
        Dim messageBoxText As String
        messageBoxText = "  " & myCulture & "     " & myDTFI.FullDateTimePattern
        MessageBoxU 0, StrPtr(messageBoxText), StrPtr("Culture FullDateTimePattern"), 0
    #End If
End Sub

'/*
'This code produces the following output.  The question marks take the place of native script characters.
'
' CULTURE    PROPERTY VALUE
'  en-US     dddd, MMMM dd, yyyy h:mm:ss tt
'  ja-JP     yyyy'年'M'月'd'日' H:mm:ss
'  fr-FR     dddd d MMMM yyyy HH:mm:ss
'
'*/
