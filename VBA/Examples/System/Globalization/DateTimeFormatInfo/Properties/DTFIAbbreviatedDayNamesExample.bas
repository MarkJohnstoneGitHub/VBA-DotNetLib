Attribute VB_Name = "DTFIAbbreviatedDayNamesExample"
'@Folder "Examples.System.Globalization.DateTimeFormatInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 1, 2023
'@LastModified September 1, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.abbreviateddaynames?view=netframework-4.8.1#examples

Option Explicit

' The following example creates a read/write CultureInfo object that represents
' the English (United States) culture and assigns abbreviated day names to its
' AbbreviatedDayNames property. It then uses the "ddd" format specifier in a
' custom date and time format string to display the string representation of
' dates for one week beginning May 28, 2014.
Public Sub DateTimeFormatInfoAbbreviatedDayNames()
    Dim ci As DotNetLib.CultureInfo
    Set ci = CultureInfo.CreateSpecificCulture("en-US")
    Dim dtfi As DotNetLib.DateTimeFormatInfo
    Set dtfi = ci.DateTimeFormat
    dtfi.SetAbbreviatedDayNames StringArray.CreateInitialize1D("Su", "M", "Tu", "W", _
                                                "Th", "F", "Sa")
                                                
    Dim dat As DotNetLib.DateTime
    Set dat = DateTime.CreateFromDate(2014, 5, 28)
    
    Dim ctr As Long
    For ctr = 0 To 6
        Dim Output As String
        Output = dat.AddDays(ctr).ToString2("ddd MMM dd, yyyy", ci)
        Debug.Print Output
    Next
End Sub

' The example displays the following output:
'       W May 28, 2014
'       Th May 29, 2014
'       F May 30, 2014
'       Sa May 31, 2014
'       Su Jun 01, 2014
'       M Jun 02, 2014
'       Tu Jun 03, 2014


