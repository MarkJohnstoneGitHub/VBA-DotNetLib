Attribute VB_Name = "DTFIAbbreviatedMonthNamesExample"
'@Folder("Examples.System.Globalization.DateTimeFormatInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 1, 2023
'@LastModified September 1, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.abbreviatedmonthnames?view=netframework-4.8.1#examples
' https://stackoverflow.com/questions/13185159/how-to-pass-byte-arrays-as-udt-properties-from-vb6-vba-to-c-sharp-com-dll
' https://stackoverflow.com/a/13185863/10759363

Option Explicit

' The following example creates a read/write CultureInfo object that represents
' the English (United States) culture and assigns abbreviated genitive month names
' to its AbbreviatedMonthNames and AbbreviatedMonthGenitiveNames properties.
' It then displays the string representation of dates that include the abbreviated
' name of each month in the culture's supported calendar.
Public Sub DateTimeFormatInfoAbbreviatedMonthNames()
    Dim ci As DotNetLib.CultureInfo
    Set ci = CultureInfo.CreateSpecificCulture("en-US")
    Dim dtfi As DotNetLib.DateTimeFormatInfo
    Set dtfi = ci.DateTimeFormat
    
    dtfi.SetAbbreviatedMonthNames Strings.ToArray("of Jan", "of Feb", "of Mar", _
                                                "of Apr", "of May", "of Jun", _
                                                "of Jul", "of Aug", "of Sep", _
                                                "of Oct", "of Nov", "of Dec", "")
                                                
    dtfi.SetAbbreviatedMonthGenitiveNames dtfi.AbbreviatedMonthNames
    Dim dat As DotNetLib.DateTime
    Set dat = DateTime.CreateFromDate(2012, 5, 28)
    
    Dim ctr As Long
    For ctr = 0 To dtfi.Calendar.GetMonthsInYear(dat.Year) - 1
        Debug.Print dat.AddMonths(ctr).ToString4("dd MMM yyyy", dtfi)
    Next
End Sub

' The example displays the following output:
'       28 of May 2012
'       28 of Jun 2012
'       28 of Jul 2012
'       28 of Aug 2012
'       28 of Sep 2012
'       28 of Oct 2012
'       28 of Nov 2012
'       28 of Dec 2012
'       28 of Jan 2013
'       28 of Feb 2013
'       28 of Mar 2013
'       28 of Apr 2013
