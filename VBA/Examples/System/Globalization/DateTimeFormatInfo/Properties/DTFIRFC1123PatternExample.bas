Attribute VB_Name = "DTFIRFC1123PatternExample"
'@Folder("Examples.System.Globalization.DateTimeFormatInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 2, 2023
'@LastModified September 2, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.rfc1123pattern?view=netframework-4.8.1#examples

Option Explicit

' The following example displays the value of LongTimePattern for a few cultures.
Public Sub DateTimeFormatInfoRFC1123Pattern()
    ' Displays the values of the pattern properties.
    Debug.Print " CULTURE    PROPERTY VALUE"
    PrintPattern "en-US"
    PrintPattern "ja-JP"
    PrintPattern "fr-FR"
End Sub

Private Sub PrintPattern(ByVal myCulture As String)
    Dim myDTFI As DotNetLib.DateTimeFormatInfo
    Set myDTFI = CultureInfo.CreateFromName(myCulture, False).DateTimeFormat
    Debug.Print "  "; myCulture; "     "; myDTFI.RFC1123Pattern
End Sub

'/*
'This code produces the following output.
'
' CULTURE    PROPERTY VALUE
'  en-US     ddd, dd MMM yyyy HH':'mm':'ss 'GMT'
'  ja-JP     ddd, dd MMM yyyy HH':'mm':'ss 'GMT'
'  fr-FR     ddd, dd MMM yyyy HH':'mm':'ss 'GMT'
'
'*/
