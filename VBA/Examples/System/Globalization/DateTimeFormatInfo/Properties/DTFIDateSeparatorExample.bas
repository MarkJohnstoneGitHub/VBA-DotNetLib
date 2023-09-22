Attribute VB_Name = "DTFIDateSeparatorExample"
'@Folder("Examples.System.Globalization.DateTimeFormatInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 1, 2023
'@LastModified September 9, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.dateseparator?view=netframework-4.8.1#examples

Option Explicit

' The following example instantiates a CultureInfo object for the en-US culture,
' changes its date separator to "-", and displays a date by using the
' "d", "G", and "g" standard format strings.
Public Sub DateTimeFormatInfoDateSeparator()
    Dim value As DotNetLib.DateTime
    Set value = DateTime.CreateFromDate(2013, 9, 8)
    
    Dim formats() As String
    formats = StringArray.ToArray("d", "G", "g")
    Dim culture As DotNetLib.CultureInfo
    Set culture = CultureInfo.CreateSpecificCulture("en-US")
    Dim dtfi As DotNetLib.DateTimeFormatInfo
    Set dtfi = culture.DateTimeFormat
    dtfi.DateSeparator = "-"
    
    Dim fmt As Variant
    For Each fmt In formats
        Debug.Print fmt; ": "; value.ToString2(fmt, dtfi)
    Next
End Sub

' The example displays the following output:
'       d: 9-8-2013
'       G: 9-8-2013 12:00:00 AM
'       g: 9-8-2013 12:00 AM
