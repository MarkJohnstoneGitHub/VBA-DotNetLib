Attribute VB_Name = "DateTimeGetDateTimeFormats3Eg"
'@Folder("Examples.System.DateTime.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 2, 2023
'@LastModified September 9, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.getdatetimeformats?view=netframework-4.8.1#system-datetime-getdatetimeformats(system-iformatprovider)

Option Explicit

' The following example demonstrates the DateTime.GetDateTimeFormats(IFormatProvider) method.
' It displays the string representation of a date using all possible standard date and time
' formats for the fr-FR culture.
Public Sub DateTimeGetDateTimeFormats3()
    Dim july28 As DotNetLib.DateTime
    Set july28 = DateTime.CreateFromDateTime(2009, 7, 28, 5, 23, 15, 16)
   
    Dim culture As DotNetLib.CultureInfo  ' Used DotNetLib.CultureInfo instead of IFormatProvider so can inspect in locals window.
    Set culture = CultureInfo.CreateFromName("fr-FR", True)
   
    ' Get the short date formats using the "fr-FR" culture.
    Dim frenchJuly28Formats() As String
    frenchJuly28Formats = july28.GetDateTimeFormats3(culture)

    ' Display july28 in various formats using "fr-FR" culture.
    Dim varFormat As Variant
    For Each varFormat In frenchJuly28Formats
       Debug.Print varFormat
    Next
End Sub

'The example displays the following output:
'28/07/2009
'28/07/09
'28.07.09
'28-07-09
'2009-07-28
'mardi 28 juillet 2009
'28 juil. 09
'28 juillet 2009
'mardi 28 juillet 2009 05:23
'mardi 28 juillet 2009 5:23
'mardi 28 juillet 2009 05.23
'mardi 28 juillet 2009 05 h 23
'28 juil. 09 05:23
'28 juil. 09 5:23
'28 juil. 09 05.23
'28 juil. 09 05 h 23
'28 juillet 2009 05:23
'28 juillet 2009 5:23
'28 juillet 2009 05.23
'28 juillet 2009 05 h 23
'mardi 28 juillet 2009 05:23:15
'mardi 28 juillet 2009 5:23:15
'mardi 28 juillet 2009 05.23
'mardi 28 juillet 2009 05 h 23
'28 juil. 09 05:23:15
'28 juil. 09 5:23:15
'28 juil. 09 05.23
'28 juil. 09 05 h 23
'28 juillet 2009 05:23:15
'28 juillet 2009 5:23:15
'28 juillet 2009 05.23
'28 juillet 2009 05 h 23
'28/07/2009 05:23
'28/07/2009 5:23
'28/07/2009 05.23
'28/07/2009 05 h 23
'28/07/09 05:23
'28/07/09 5:23
'28/07/09 05.23
'28/07/09 05 h 23
'28.07.09 05:23
'28.07.09 5:23
'28.07.09 05.23
'28.07.09 05 h 23
'28-07-09 05:23
'28-07-09 5:23
'28-07-09 05.23
'28-07-09 05 h 23
'2009-07-28 05:23
'2009-07-28 5:23
'2009-07-28 05.23
'2009-07-28 05 h 23
'28/07/2009 05:23:15
'28/07/2009 5:23:15
'28/07/2009 05.23
'28/07/2009 05 h 23
'28/07/09 05:23:15
'28/07/09 5:23:15
'28/07/09 05.23
'28/07/09 05 h 23
'28.07.09 05:23:15
'28.07.09 5:23:15
'28.07.09 05.23
'28.07.09 05 h 23
'28-07-09 05:23:15
'28-07-09 5:23:15
'28-07-09 05.23
'28-07-09 05 h 23
'2009-07-28 05:23:15
'2009-07-28 5:23:15
'2009-07-28 05.23
'2009-07-28 05 h 23
'28 juillet
'28 juillet
'2009-07-28T05:23:15.0160000
'2009-07-28T05:23:15.0160000
'Tue, 28 Jul 2009 05:23:15 GMT
'Tue, 28 Jul 2009 05:23:15 GMT
'2009-07-28T05:23:15
'05:23
'5:23
'05.23
'5  h 23
'05:23:15
'5:23:15
'05.23
'5  h 23
'2009-07-28 05:23:15Z
'mardi 28 juillet 2009 12:23:15
'mardi 28 juillet 2009 12:23:15
'mardi 28 juillet 2009 12.23
'mardi 28 juillet 2009 12 h 23
'28 juil. 09 12:23:15
'28 juil. 09 12:23:15
'28 juil. 09 12.23
'28 juil. 09 12 h 23
'28 juillet 2009 12:23:15
'28 juillet 2009 12:23:15
'28 juillet 2009 12.23
'28 juillet 2009 12 h 23
'juillet 2009
'juillet 2009
