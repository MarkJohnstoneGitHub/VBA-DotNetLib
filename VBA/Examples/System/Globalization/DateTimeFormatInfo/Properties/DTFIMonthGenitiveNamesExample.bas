Attribute VB_Name = "DTFIMonthGenitiveNamesExample"
'@Folder "Examples.System.Globalization.DateTimeFormatInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 2, 2023
'@LastModified September 2, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.monthgenitivenames?view=netframework-4.8.1#examples

Option Explicit

' The following example demonstrates several methods and properties that specify
' date and time format patterns, native calendar name, and full and abbreviated
' month and day names.
'
' This code example demonstrates the DateTimeFormatInfo
' MonthGenitiveNames, AbbreviatedMonthGenitiveNames,
' ShortestDayNames, and NativeCalendarName properties, and
' the GetShortestDayName() and SetAllDateTimePatterns() methods.
Public Sub DateTimeFormatInfoMonthGenitiveNames()
    Dim myDateTimePatterns() As String
    myDateTimePatterns = StringArray.CreateInitialize1D("MM/dd/yy", "MM/dd/yyyy")
    
    ' Get the en-US culture.
    Dim ci As DotNetLib.CultureInfo
    Set ci = CultureInfo.CreateFromName("en-US")
    ' Get the DateTimeFormatInfo for the en-US culture.
    Dim dtfi As DotNetLib.DateTimeFormatInfo
    Set dtfi = ci.DateTimeFormat
    
    ' Display the effective culture.
    Debug.Print "This code example uses the "; ci.name; " culture."
    
    ' Display the native calendar name.
    Debug.Print VBA.vbNewLine; "MonthGenitiveNames..."
    Debug.Print """"; dtfi.NativeCalendarName; """"
    
    ' Display month genitive names.
    Debug.Print VBA.vbNewLine; "NativeCalendarName..."
    Dim varName As Variant
    For Each varName In dtfi.MonthGenitiveNames
        Debug.Print """"; varName; """"
    Next
    
    ' Display abbreviated month genitive names.
    Debug.Print VBA.vbNewLine; "AbbreviatedMonthGenitiveNames..."
    For Each varName In dtfi.AbbreviatedMonthGenitiveNames
        Debug.Print """"; varName; """"
    Next
        
    ' Display shortest day names.
    Debug.Print VBA.vbNewLine; "ShortestDayNames..."
    For Each varName In dtfi.ShortestDayNames
        Debug.Print """"; varName; """"
    Next

    ' Display shortest day name for a particular day of the week.
    Debug.Print VBA.vbNewLine; "GetShortestDayName(DayOfWeek.Sunday)..."
    Debug.Print """"; dtfi.GetShortestDayName(DayOfWeek.DayOfWeek_Sunday); """"

    ' Display the initial DateTime format patterns for the 'd' format specifier.
    Debug.Print VBA.vbNewLine; "Initial DateTime format patterns for the 'd' format specifier..."
    For Each varName In dtfi.GetAllDateTimePatterns("d")
        Debug.Print """"; varName; """"
    Next

    ' Change the initial DateTime format patterns for the 'd' DateTime format specifier.
    Debug.Print VBA.vbNewLine; "Change the initial DateTime format patterns for the "; VBA.vbNewLine; _
                "'d' format specifier to my format patterns..."
    
    dtfi.SetAllDateTimePatterns myDateTimePatterns, "d"

    ' Display the new DateTime format patterns for the 'd' format specifier.
    Debug.Print VBA.vbNewLine; "New DateTime format patterns for the 'd' format specifier..."
    For Each varName In dtfi.GetAllDateTimePatterns("d")
        Debug.Print """"; varName; """"
    Next
End Sub

'/*
'This code example produces the following results:
'
'This code example uses the en-US culture.
'
'NativeCalendarName...
'"Gregorian Calendar"
'
'MonthGenitiveNames...
'"January"
'"February"
'"March"
'"April"
'"May"
'"June"
'"July"
'"August"
'"September"
'"October"
'"November"
'"December"
'""
'
'AbbreviatedMonthGenitiveNames...
'"Jan"
'"Feb"
'"Mar"
'"Apr"
'"May"
'"Jun"
'"Jul"
'"Aug"
'"Sep"
'"Oct"
'"Nov"
'"Dec"
'""
'
'ShortestDayNames...
'"Su"
'"Mo"
'"Tu"
'"We"
'"Th"
'"Fr"
'"Sa"
'
'GetShortestDayName(DayOfWeek.Sunday)...
'"Su"
'
'Initial DateTime format patterns for the 'd' format specifier...
'"M/d/yyyy"
'"M/d/yy"
'"MM/dd/yy"
'"MM/dd/yyyy"
'"yy/MM/dd"
'"yyyy-MM-dd"
'"dd-MMM-yy"
'
'Change the initial DateTime format patterns for the
''d' format specifier to my format patterns...
'
'New DateTime format patterns for the 'd' format specifier...
'"MM/dd/yy"
'"MM/dd/yyyy"
'
'*/


