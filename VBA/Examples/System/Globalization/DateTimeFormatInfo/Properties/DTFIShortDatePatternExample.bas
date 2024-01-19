Attribute VB_Name = "DTFIShortDatePatternExample"
'@Folder "Examples.System.Globalization.DateTimeFormatInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 2, 2023
'@LastModified January 6, 2024

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.shortdatepattern?view=netframework-4.8.1#examples

Option Explicit

' The following example displays the value of the ShortDatePattern property and
' the value of a date formatted using the ShortDatePattern property for a few cultures.
Public Sub DateTimeFormatInfoShortDatePattern()
    Dim cultures() As String
    cultures = StringArray.CreateInitialize1D("en-US", "ja-JP", "fr-FR")
    Dim date1 As DotNetLib.DateTime
    Set date1 = DateTime.CreateFromDate(2011, 5, 1)
    
    Debug.Print VBString.Format(VBString.Unescape(" {0,7} {1,19} {2,10}\n"), "CULTURE", "PROPERTY VALUE", "DATE")

    Dim culture As Variant
    For Each culture In cultures
        Dim dtfi As DotNetLib.DateTimeFormatInfo
        Set dtfi = CultureInfo.CreateSpecificCulture(culture).DateTimeFormat
        Debug.Print VBString.Format(" {0,7} {1,19} {2,10}", culture, _
                           dtfi.ShortDatePattern, _
                           date1.ToString2("d", dtfi))
    Next
End Sub

' The example displays the following output:
'        CULTURE      PROPERTY VALUE       DATE
'
'          en-US            M/d/yyyy   5/1/2011
'          ja-JP          yyyy/MM/dd 2011/05/01
'          fr-FR          dd/MM/yyyy 01/05/2011


