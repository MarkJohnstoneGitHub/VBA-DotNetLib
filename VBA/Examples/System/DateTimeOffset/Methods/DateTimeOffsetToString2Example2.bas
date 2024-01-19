Attribute VB_Name = "DateTimeOffsetToString2Example2"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 26, 2023
'@LastModified January 10, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.tostring?view=netframework-4.8.1#system-datetimeoffset-tostring(system-string-system-iformatprovider)

Option Explicit

''
' The following example uses the ToString(String, IFormatProvider) method to
' display a DateTimeOffset object using a custom format string for several
' different cultures.
''
Public Sub DateTimeOffsetToString2()
    Dim outputDate As DotNetLib.DateTimeOffset
    Set outputDate = DateTimeOffset.CreateFromDateTimeParts(2007, 11, 1, 9, 0, 0, TimeSpan.Create(-7, 0, 0))
    Dim pvtFormat As String
    pvtFormat = "dddd, MMM dd yyyy HH:mm:ss zzz"
    
    ' Output date and time using custom format specification
    Debug.Print outputDate.ToString2(pvtFormat, Nothing)
    Debug.Print outputDate.ToString2(pvtFormat, CultureInfo.InvariantCulture)
    Debug.Print outputDate.ToString2(pvtFormat, CultureInfo.CreateFromName("fr-FR"))
    Debug.Print outputDate.ToString2(pvtFormat, CultureInfo.CreateFromName("es-ES"))
End Sub

' The example displays the following output to the console:
'    Thursday, Nov 01 2007 09:00:00 -07:00
'    Thursday, Nov 01 2007 09:00:00 -07:00
'    jeudi, nov. 01 2007 09:00:00 -07:00
'    jueves, nov 01 2007 09:00:00 -07:00
