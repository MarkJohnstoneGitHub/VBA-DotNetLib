Attribute VB_Name = "DateTimeGetDateTimeFormats2Eg"
'@Folder("Examples.System.DateTime.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 3, 2023
'@LastModified September 9, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.getdatetimeformats?view=netframework-4.8.1#system-datetime-getdatetimeformats(system-char-system-iformatprovider)

Option Explicit

' The following example demonstrates the GetDateTimeFormats(Char, IFormatProvider) method.
' It displays the string representations of a date using the short date format specifier
' ("d") for the fr-FR culture.
Public Sub DateTimeGetDateTimeFormats2()
    Dim july28 As DotNetLib.DateTime
    Set july28 = DateTime.CreateFromDateTime(2009, 7, 28, 5, 23, 15, 16)
   
    Dim culture As DotNetLib.CultureInfo  ' Used DotNetLib.CultureInfo instead of IFormatProvider so can inspect in locals window.
    Set culture = CultureInfo.CreateFromName("fr-FR", True)
   
    ' Get the short date formats using the "fr-FR" culture.
    Dim frenchJuly28Formats() As String
    frenchJuly28Formats = july28.GetDateTimeFormats2("d", culture)

    ' Display july28 in various formats using "fr-FR" culture.
    Dim varFormat As Variant
    For Each varFormat In frenchJuly28Formats
       Debug.Print varFormat
    Next
End Sub

' The example displays the following output:
'       28/07/2009
'       28/07/09
'       28.07.09
'       28-07-09
'       2009-07-28
