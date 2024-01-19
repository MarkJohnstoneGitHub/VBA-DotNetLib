Attribute VB_Name = "DTFITimeSeparatorExample"
'@Folder "Examples.System.Globalization.DateTimeFormatInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 4, 2023
'@LastModified September 9, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.timeseparator?view=netframework-4.8.1#examples
Option Explicit

' The following example instantiates a CultureInfo object for the en-US culture,
' changes its date separator to ".", and displays a date by using the "t", "T",
' "F", "f", "G", and "g" standard format strings.
Public Sub DateTimeFormatInfoTimeSeparator()
    Dim value As DotNetLib.DateTime
    Set value = DateTime.CreateFromDateTime(2013, 9, 8, 14, 30, 0)
    
    Dim formats() As String
    formats = StringArray.CreateInitialize1D("t", "T", "f", "F", "G", "g")
    Dim culture As DotNetLib.CultureInfo
    Set culture = CultureInfo.CreateSpecificCulture("en-US")
    Dim dtfi As DotNetLib.DateTimeFormatInfo
    Set dtfi = culture.DateTimeFormat
    dtfi.TimeSeparator = "."
    
    Dim fmt As Variant
    For Each fmt In formats
        Debug.Print fmt; ": "; value.ToString2(fmt, dtfi)
    Next
    
End Sub

' The example displays the following output:
'       t: 2.30 PM
'       T: 2.30.00 PM
'       f: Sunday, September 08, 2013 2.30 PM
'       F: Sunday, September 08, 2013 2.30.00 PM
'       G: 9/8/2013 2.30.00 PM
'       g: 9/8/2013 2.30 PM
