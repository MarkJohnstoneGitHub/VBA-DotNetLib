Attribute VB_Name = "DTFISetAllDateTimePatternsEg"
'@Folder("Examples.System.Globalization.DateTimeFormatInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 5, 2023
'@LastModified September 5, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.setalldatetimepatterns?view=netframework-4.8.1#examples

Option Explicit

' The following example instantiates a CultureInfo object that represents the
' "en-US" (English - United States) culture and uses it to parse an array of
' date and time strings using the "Y" standard format string. It then uses the
' SetAllDateTimePatterns method to associate a new custom format string with
' the "Y" standard format string, and then attempts to parse the array of date
' and time strings.
' Output from the example demonstrates that the new custom format string is
' used in both the parsing and formatting operations.
Public Sub DateTimeFormatInfoSetAllDateTimePatterns()
    ' Use standard en-US culture.
    Dim enUS As DotNetLib.CultureInfo
    Set enUS = CultureInfo.CreateFromName("en-US")
    
    Dim values() As String
    values = StringArray.ToArray("December 2010", "December, 2010", _
                            "Dec-2010", "December-2010")
    
    Debug.Print "Supported Y/y patterns for "; enUS.name; " culture:"
    
    Dim pattern As Variant
    For Each pattern In enUS.DateTimeFormat.GetAllDateTimePatterns("Y")
        Debug.Print "   " + pattern
    Next
    Debug.Print
    
    ' Try to parse each date string using "Y" format specifier.
    Dim value  As Variant
    For Each value In values
        Dim dat As DotNetLib.DateTime
        On Error Resume Next
        Set dat = DateTime.ParseExact(value, "Y", enUS)
        If Try Then
            Debug.Print "   Parsed "; value; " as "; dat.ToString2("Y")
        ElseIf Catch(FormatException) Then
            Debug.Print "   Cannot parse "; value
        End If
        On Error GoTo 0 'reset error handling
    Next
    Debug.Print
    
    'Modify supported "Y" format.
    enUS.DateTimeFormat.SetAllDateTimePatterns StringArray.ToArray("MMM-yyyy"), "Y"
    Debug.Print "Supported Y/y patterns for "; enUS.name; " culture:"
    For Each pattern In enUS.DateTimeFormat.GetAllDateTimePatterns("Y")
        Debug.Print "   " + pattern
    Next
    Debug.Print

    ' Try to parse each date string using "Y" format specifier.
    For Each value In values
        On Error Resume Next
        Set dat = DateTime.ParseExact(value, "Y", enUS)
        If Try Then
            Debug.Print "   Parsed "; value; " as "; dat.ToString2("Y")
        ElseIf Catch(FormatException) Then
            Debug.Print "   Cannot parse "; value
        End If
        On Error GoTo 0 'reset error handling
    Next
End Sub

' The example displays the following output:
'       Supported Y/y patterns for en-US culture:
'          MMMM, yyyy
'
'          Cannot parse December 2010
'          Parsed December, 2010 as December, 2010
'          Cannot parse Dec-2010
'          Cannot parse December-2010
'
'       New supported Y/y patterns for en-US culture:
'          MMM-yyyy
'
'          Cannot parse December 2010
'          Cannot parse December, 2010
'          Parsed Dec-2010 as Dec-2010
'          Cannot parse December-2010

