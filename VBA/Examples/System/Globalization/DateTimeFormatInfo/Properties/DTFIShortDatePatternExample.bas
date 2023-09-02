Attribute VB_Name = "DTFIShortDatePatternExample"
'@Folder("Examples.System.Globalization.DateTimeFormatInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 2, 2023
'@LastModified September 2, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.shortdatepattern?view=netframework-4.8.1#examples

Option Explicit

' The following example displays the value of the ShortDatePattern property and
' the value of a date formatted using the ShortDatePattern property for a few cultures.
Public Sub DateTimeFormatInfoShortDatePattern()
    Dim cultures() As String
    cultures = Strings.ToArray("en-US", "ja-JP", "fr-FR")
    Dim date1 As DotNetLib.DateTime
    Set date1 = DateTime.CreateFromDate(2011, 5, 1)
    
    Debug.Print Align("CULTURE", 7, Justify_Right); "  "; Align("PROPERTY VALUE", 19, Justify_Right); "  "; Align("DATE", 10, Justify_Right)

    Dim culture As Variant
    For Each culture In cultures
        Dim dtfi As DotNetLib.DateTimeFormatInfo
        Set dtfi = CultureInfo.CreateSpecificCulture(culture).DateTimeFormat
        Debug.Print Align(culture, 7, Justify_Right); "  ";
        Debug.Print Align(dtfi.ShortDatePattern, 19, Justify_Right); "  ";
        Debug.Print Align(date1.ToString4("d", dtfi), 10, Justify_Right)
    Next
End Sub

' The example displays the following output:
'        CULTURE      PROPERTY VALUE       DATE
'
'          en-US            M/d/yyyy   5/1/2011
'          ja-JP          yyyy/MM/dd 2011/05/01
'          fr-FR          dd/MM/yyyy 01/05/2011

