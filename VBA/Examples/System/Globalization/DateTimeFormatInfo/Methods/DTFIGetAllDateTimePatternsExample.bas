Attribute VB_Name = "DTFIGetAllDateTimePatternsExample"
'@Folder("Examples.System.Globalization.DateTimeFormatInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 4, 2023
'@LastModified September 4, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.getalldatetimepatterns?view=netframework-4.8.1#system-globalization-datetimeformatinfo-getalldatetimepatterns

Option Explicit

' The following example displays the date and time format strings for the invariant culture,
' as well as the result string that is produced when that format string is used to format
' a particular date.
Public Sub DateTimeFormatInfoGetAllDateTimePatterns()
    Dim date1 As DotNetLib.DateTime
    Set date1 = DateTime.CreateFromDateTime(2014, 8, 28, 12, 28, 30)
    Dim invDTF As DotNetLib.DateTimeFormatInfo
    Set invDTF = DateTimeFormatInfo.Create()
    Dim formats() As String
    formats = invDTF.GetAllDateTimePatterns()
    
    Debug.Print Align("Pattern", 40, Justify_Left); " "; "Result String"; VBA.vbNewLine
    Dim fmt As Variant
    For Each fmt In formats
       Debug.Print Align(fmt, 40, Justify_Left); " "; date1.ToString2(fmt)
    Next
   
End Sub

' The example displays the following output:
'    Pattern                                  Result String
'
'    MM/dd/yyyy                               08/28/2014
'    yyyy-MM-dd                               2014-08-28
'    dddd, dd MMMM yyyy                       Thursday, 28 August 2014
'    dddd, dd MMMM yyyy HH:mm                 Thursday, 28 August 2014 12:28
'    dddd, dd MMMM yyyy hh:mm tt              Thursday, 28 August 2014 12:28 PM
'    dddd, dd MMMM yyyy H:mm                  Thursday, 28 August 2014 12:28
'    dddd, dd MMMM yyyy h:mm tt               Thursday, 28 August 2014 12:28 PM
'    dddd, dd MMMM yyyy HH:mm:ss              Thursday, 28 August 2014 12:28:30
'    MM/dd/yyyy HH:mm                         08/28/2014 12:28
'    MM/dd/yyyy hh:mm tt                      08/28/2014 12:28 PM
'    MM/dd/yyyy H:mm                          08/28/2014 12:28
'    MM/dd/yyyy h:mm tt                       08/28/2014 12:28 PM
'    yyyy-MM-dd HH:mm                         2014-08-28 12:28
'    yyyy-MM-dd hh:mm tt                      2014-08-28 12:28 PM
'    yyyy-MM-dd H:mm                          2014-08-28 12:28
'    yyyy-MM-dd h:mm tt                       2014-08-28 12:28 PM
'    MM/dd/yyyy HH:mm:ss                      08/28/2014 12:28:30
'    yyyy-MM-dd HH:mm:ss                      2014-08-28 12:28:30
'    MMMM dd                                  August 28
'    MMMM dd                                  August 28
'    yyyy'-'MM'-'dd'T'HH':'mm':'ss.fffffffK   2014-08-28T12:28:30.0000000
'    yyyy'-'MM'-'dd'T'HH':'mm':'ss.fffffffK   2014-08-28T12:28:30.0000000
'    ddd, dd MMM yyyy HH':'mm':'ss 'GMT'      Thu, 28 Aug 2014 12:28:30 GMT
'    ddd, dd MMM yyyy HH':'mm':'ss 'GMT'      Thu, 28 Aug 2014 12:28:30 GMT
'    yyyy'-'MM'-'dd'T'HH':'mm':'ss            2014-08-28T12:28:30
'    HH:mm                                    12:28
'    hh:mm tt                                 12:28 PM
'    H:mm                                     12:28
'    h:mm tt                                  12:28 PM
'    HH:mm:ss                                 12:28:30
'    yyyy'-'MM'-'dd HH':'mm':'ss'Z'           2014-08-28 12:28:30Z
'    dddd, dd MMMM yyyy HH:mm:ss              Thursday, 28 August 2014 12:28:30
'    yyyy MMMM                                2014 August
'    yyyy MMMM                                2014 August
