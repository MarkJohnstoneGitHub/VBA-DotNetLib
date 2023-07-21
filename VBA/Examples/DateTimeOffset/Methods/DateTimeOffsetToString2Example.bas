Attribute VB_Name = "DateTimeOffsetToString2Example"
'@Folder("VBADotNetLib.Examples.DateTimeOffset.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified July 22, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.tostring?view=netframework-4.8.1#system-datetimeoffset-tostring(system-string)

Option Explicit

'@Description("The following example displays a DateTimeOffset object to the console using each of the standard date and time format specifiers.")
' The output is formatted by using the en-us culture.
Public Sub DateTimeOffsetToString2()
Attribute DateTimeOffsetToString2.VB_Description = "The following example displays a DateTimeOffset object to the console using each of the standard date and time format specifiers."
    Dim outputDate As DateTimeOffset
    Set outputDate = DateTimeOffset.CreateFromDateTimeParts(2007, 10, 31, 21, 0, 0, TimeSpan.Create(-8, 0, 0))
    Dim specifier As String
    
    '// Output date using each standard date/time format specifier
    specifier = "d"
    ' Displays   d: 10/31/2007
    Debug.Print specifier & ": " & outputDate.ToString2(specifier)
    
    specifier = "D"
    ' Displays   D: Wednesday, October 31, 2007
    Debug.Print specifier & ": " & outputDate.ToString2(specifier)
    
    specifier = "t"
    ' Displays   t: 9:00 PM
    Debug.Print specifier & ": " & outputDate.ToString2(specifier)
    
    specifier = "T"
    ' Displays   T: 9:00:00 PM
    Debug.Print specifier & ": " & outputDate.ToString2(specifier)
    
    specifier = "f"
    ' Displays   f: Wednesday, October 31, 2007 9:00 PM
    Debug.Print specifier & ": " & outputDate.ToString2(specifier)
    
    specifier = "F"
    ' Displays   F: Wednesday, October 31, 2007 9:00:00 PM
    Debug.Print specifier & ": " & outputDate.ToString2(specifier)

    specifier = "g"
    ' Displays   g: 10/31/2007 9:00 PM
    Debug.Print specifier & ": " & outputDate.ToString2(specifier)

    specifier = "G"
    ' Displays   G: 10/31/2007 9:00:00 PM
    Debug.Print specifier & ": " & outputDate.ToString2(specifier)
    
    specifier = "M"           ' 'm' is identical
    ' Displays   M: October 31
    Debug.Print specifier & ": " & outputDate.ToString2(specifier)

    specifier = "R"           ' 'r' is identical
    ' Displays   R: Thu, 01 Nov 2007 05:00:00 GMT
    Debug.Print specifier & ": " & outputDate.ToString2(specifier)

    specifier = "s"
    ' Displays   s: 2007-10-31T21:00:00
    Debug.Print specifier & ": " & outputDate.ToString2(specifier)
    
    specifier = "u"
    ' Displays   u: 2007-11-01 05:00:00Z
    Debug.Print specifier & ": " & outputDate.ToString2(specifier)

    ' Specifier is not supported
    specifier = "U"
    On Error Resume Next
    Debug.Print specifier & ": " & outputDate.ToString2(specifier)
    If Catch(FormatException) Then
       Debug.Print specifier & ": Not supported."
    End If
    On Error GoTo 0 'Stop code and display error
   
    specifier = "Y"         ' 'y' is identical
    ' Displays   Y: October, 2007
    Debug.Print specifier & ": " & outputDate.ToString2(specifier)
End Sub

' The output is formatted by using the en-us culture.
'd: 10/31/2007
'D: Wednesday, October 31, 2007
't: 9:00 PM
'T: 9:00:00 PM
'f: Wednesday, October 31, 2007 9:00 PM
'F: Wednesday, October 31, 2007 9:00:00 PM
'g: 10/31/2007 9:00 PM
'G: 10/31/2007 9:00:00 PM
'M: October 31
'R: Thu, 01 Nov 2007 05:00:00 GMT
's: 2007-10-31T21:00:00
'u: 2007-11-01 05:00:00Z
'U: Not supported.
'Y: October 2007
