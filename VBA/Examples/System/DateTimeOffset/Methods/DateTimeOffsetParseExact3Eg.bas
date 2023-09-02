Attribute VB_Name = "DateTimeOffsetParseExact3Eg"
'@Folder("Examples.System.DateTimeOffset.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 27, 2023
'@LastModified August 27, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.parseexact?view=netframework-4.8.1#system-datetimeoffset-parseexact(system-string-system-string()-system-iformatprovider-system-globalization-datetimestyles)

'@TODO Implement DateTimeFormat in DotNetLib

Option Explicit

' The following example defines multiple input formats for the string
' representation of a date and time and offset value, and then passes
' the string that is entered by the user to the
' DateTimeOffset.ParseExact(String, String[], IFormatProvider, DateTimeStyles) method.
Public Sub DateTimeOffsetParseExact3()
    Dim tries As Long
    Dim pvtInput As String
    
    Dim formats() As String
    formats = Strings.ToArray( _
                "@M/dd/yyyy HH:m zzz", "MM/dd/yyyy HH:m zzz", _
                "M/d/yyyy HH:m zzz", "MM/d/yyyy HH:m zzz", _
                "M/dd/yy HH:m zzz", "MM/dd/yy HH:m zzz", _
                "M/d/yy HH:m zzz", "MM/d/yy HH:m zzz", _
                "M/dd/yyyy H:m zzz", "MM/dd/yyyy H:m zzz", _
                "M/d/yyyy H:m zzz", "MM/d/yyyy H:m zzz", _
                "M/dd/yy H:m zzz", "MM/dd/yy H:m zzz", _
                "M/d/yy H:m zzz", "MM/d/yy H:m zzz", _
                "M/dd/yyyy HH:mm zzz", "MM/dd/yyyy HH:mm zzz", _
                "M/d/yyyy HH:mm zzz", "MM/d/yyyy HH:mm zzz", _
                "M/dd/yy HH:mm zzz", "MM/dd/yy HH:mm zzz", _
                "M/d/yy HH:mm zzz", "MM/d/yy HH:mm zzz", _
                "M/dd/yyyy H:mm zzz", "MM/dd/yyyy H:mm zzz", _
                "M/d/yyyy H:mm zzz", "MM/d/yyyy H:mm zzz", _
                "M/dd/yy H:mm zzz", "MM/dd/yy H:mm zzz", _
                "M/d/yy H:mm zzz", "MM/d/yy H:mm zzz")
                
    Dim provider As DotNetLib.DateTimeFormatInfo
    Set provider = CultureInfo.InvariantCulture.DateTimeFormat
    Dim result As DotNetLib.DateTimeOffset
    Do
        pvtInput = InputBox("Enter a date, time, and offset (MM/DD/YYYY HH:MM +/-HH:MM),")
        On Error Resume Next
        Set result = DateTimeOffset.ParseExact3(pvtInput, formats, provider, DateTimeStyles.DateTimeStyles_AllowWhiteSpaces)
        If Try Then
            Debug.Print "'" & pvtInput & "' was converted to " & result.ToString()
        ElseIf Catch(FormatException) Then
            Debug.Print "Unable to parse "; "'"; pvtInput; "'"; "."
        End If
        On Error GoTo 0 'reset error handling
        tries = tries + 1
    Loop While (tries < 3)
End Sub

' Some successful sample interactions with the user might appear as follows:
'    Enter a date, time, and offset (MM/DD/YYYY HH:MM +/-HH:MM),
'    Then press Enter: 12/08/2007 6:54 -6:00
'
'    12/08/2007 6:54 -6:00 was converted to 12/8/2007 6:54:00 AM -06:00
'
'    Enter a date, time, and offset (MM/DD/YYYY HH:MM +/-HH:MM),
'    Then press Enter: 12/8/2007 06:54 -06:00
'
'    12/8/2007 06:54 -06:00 was converted to 12/8/2007 6:54:00 AM -06:00
'
'    Enter a date, time, and offset (MM/DD/YYYY HH:MM +/-HH:MM),
'    Then press Enter: 12/5/07 6:54 -6:00

