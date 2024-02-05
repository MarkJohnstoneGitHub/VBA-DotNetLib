Attribute VB_Name = "GetDateFromFileNameExample"
'@IgnoreModule VariableNotAssigned
'@Folder "Examples.System.Text.RegularExpressions.Regex"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 3, 2024
'@LastModified February 5, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://stackoverflow.com/questions/77921657/extract-date-from-string-variable-date-formats

Option Explicit

' Regular Expression testers
' https://regexr.com/
' http://regexstorm.net/tester

' Format         Regex Pattern
' MM_dd_yyyy     (0?[1-9]|1[0-2])_(0?[1-9]|[12]\d|30|31)_(\d{4})
' M.dd.yy        (0?[1-9]|1[0-2])[\.](0?[1-9]|[12]\d|30|31)[\.](\d{2})
' yyyyMMdd       \b(\d{4})(0[1-9]|1[0-2])(0[1-9]|[12]\d|30|31)\b
' MM -yyyy       (0?[1-9]|1[0-2]) -(\d{4})
' MMM dd yyyy    (?<!\w)\w{3}(?!\w)\s(0?[1-9]|[12]\d|30|31)\s(\d{4})
' MMMM yyyy      \w[a-zA-Z]+\s\d{4}

''
' Parses a string according to a list of Regular Expression patterns provided to
' obtain date strings to be converted to a date.  The date strings are parsed
' according to a list of date formats provided and culture and converted to a
' DateTime object.
' To obtain a VBA Date data type from a DotNetLib.DateTime object use the
' ToOADate member which returns a double type which can be assigned to a VBA Date.
''
Public Sub ExtractDateFromFileNameExample()
    'Regex patterns for obtaining the date string from the file name
    Dim regexPatterns() As String
    regexPatterns = StringArray.CreateInitialize1D("(0?[1-9]|1[0-2])_(0?[1-9]|[12]\d|30|31)_(\d{4})", _
                            "(0?[1-9]|1[0-2])[\.](0?[1-9]|[12]\d|30|31)[\.](\d{2})", _
                            "\b(\d{4})(0[1-9]|1[0-2])(0[1-9]|[12]\d|30|31)\b", _
                            "(0?[1-9]|1[0-2]) -(\d{4})", _
                            "(?<!\w)\w{3}(?!\w)\s(0?[1-9]|[12]\d|30|31)\s(\d{4})", _
                            "\w[a-zA-Z]+\s\d{4}")
                         
    ' https://learn.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings
    ' Datetime formats to be parsed
    Dim dateFormats() As String
    dateFormats = StringArray.CreateInitialize1D("MMM dd yyyy", "MMM d yyyy", "MM_dd_yyyy", _
                            "M.d.yy", "MM.d.yy", "M.dd.yy", "yyyyMMdd", _
                            "MM -yyyy", "MMMM yyyy", "MMM.dd.yy")
    
    'Culture used for parsing date en-US, change to parse other cultures
    Dim pvtFormatProvider As mscorlib.IFormatProvider
    Set pvtFormatProvider = CultureInfo.CreateFromName("en-US")
    
    'Sample file name data
    Dim fileNames() As String
    fileNames = StringArray.CreateInitialize1D( _
            "Feb 9 1995", _
            "Jan 09 1995", _
            "JAN 09 1995.txt", _
            "jan 09 1995", _
            "Jan 09 1995", _
            "Jan 09 1995 text", _
            "Whatever Jan 09 1995 text", _
            "Whatever Jan 9 1995 text", _
            "Whatever Jan 09 1995", _
            "01_09_1995", _
            "12_23_1995", _
            "01_09_1995.txt", _
            "Whatever 01_09_1995", _
            "Whatever01_09_1995", _
            "Whatever01_09_1995text", _
            "1.09.95", _
            "1.9.95", _
            "5.09.95", _
            "05.12.95", _
            "12.09.95", _
            "19950109", _
            "01 -1995", _
            "January 1995", _
            "December 1995")
    
    Dim fileName As Variant
    For Each fileName In fileNames
        Debug.Print VBString.Format("File name '{0}'", fileName)
        Dim dateString As String
        dateString = GetRegexPatternMatch(fileName, regexPatterns)
        If dateString <> VBA.vbNullString Then
            Dim dateValue As DotNetLib.DateTime
            If (DateTime.TryParseExact2(dateString, dateFormats, _
                                        pvtFormatProvider, _
                                        DateTimeStyles.DateTimeStyles_None, _
                                        dateValue)) Then
                'Displays DateTime object converted to a VBA date and formatted as "MM-dd-yyyy"
                Debug.Print VBString.Format("Converted '{0}' to {1:MM-dd-yyyy}.", dateString, VBA.CDate(dateValue.ToOADate))
            Else
                Debug.Print VBString.Format("Unable to convert '{0}' to a date.", dateString)
            End If
        Else
            Debug.Print VBString.Format("Unable to find date in file name '{0}'.", fileName)
        End If
        Debug.Print
    Next
End Sub

''
' Returns the matched value for a string according to a array of Regular Expressions.
''
Public Function GetRegexPatternMatch(ByVal inputStr As String, ByRef patterns() As String) As String
    Dim pvtMatchValue  As String
    pvtMatchValue = VBA.vbNullString
    Dim pattern As Variant
    For Each pattern In patterns
        Dim pvtMatch As DotNetLib.Match
        Set pvtMatch = Regex.Match(inputStr, CStr(pattern), RegexOptions.RegexOptions_IgnoreCase)
        If (pvtMatch.Success) Then
            Debug.Print VBString.Format("Pattern matched '{0}' ", pattern)
            Debug.Print VBString.Format("Found match '{0}' at position {1}.", pvtMatch.value, pvtMatch.index)
            pvtMatchValue = pvtMatch.value
            Exit For
        End If
    Next
    GetRegexPatternMatch = pvtMatchValue
 End Function
 
' Output:
'    File name 'JAN 09 1995.txt'
'    Pattern matched '(?<!\w)\w{3}(?!\w)\s(0?[1-9]|[12]\d|30|31)\s(\d{4})'
'    Found Match 'JAN 09 1995' at position 0.
'    Converted 'JAN 09 1995' to 01-09-1995.
'
'    File name 'jan 09 1995'
'    Pattern matched '(?<!\w)\w{3}(?!\w)\s(0?[1-9]|[12]\d|30|31)\s(\d{4})'
'    Found Match 'jan 09 1995' at position 0.
'    Converted 'jan 09 1995' to 01-09-1995.
'
'    File name 'Jan 09 1995'
'    Pattern matched '(?<!\w)\w{3}(?!\w)\s(0?[1-9]|[12]\d|30|31)\s(\d{4})'
'    Found Match 'Jan 09 1995' at position 0.
'    Converted 'Jan 09 1995' to 01-09-1995.
'
'    File name 'Jan 09 1995 text'
'    Pattern matched '(?<!\w)\w{3}(?!\w)\s(0?[1-9]|[12]\d|30|31)\s(\d{4})'
'    Found Match 'Jan 09 1995' at position 0.
'    Converted 'Jan 09 1995' to 01-09-1995.
'
'    File name 'Whatever Jan 09 1995 text'
'    Pattern matched '(?<!\w)\w{3}(?!\w)\s(0?[1-9]|[12]\d|30|31)\s(\d{4})'
'    Found Match 'Jan 09 1995' at position 9.
'    Converted 'Jan 09 1995' to 01-09-1995.
'
'    File name 'Whatever Jan 9 1995 text'
'    Pattern matched '(?<!\w)\w{3}(?!\w)\s(0?[1-9]|[12]\d|30|31)\s(\d{4})'
'    Found Match 'Jan 9 1995' at position 9.
'    Converted 'Jan 9 1995' to 01-09-1995.
'
'    File name 'Whatever Jan 09 1995'
'    Pattern matched '(?<!\w)\w{3}(?!\w)\s(0?[1-9]|[12]\d|30|31)\s(\d{4})'
'    Found Match 'Jan 09 1995' at position 9.
'    Converted 'Jan 09 1995' to 01-09-1995.
'
'    File name '01_09_1995'
'    Pattern matched '(0?[1-9]|1[0-2])_(0?[1-9]|[12]\d|30|31)_(\d{4})'
'    Found Match '01_09_1995' at position 0.
'    Converted '01_09_1995' to 01-09-1995.
'
'    File name '12_23_1995'
'    Pattern matched '(0?[1-9]|1[0-2])_(0?[1-9]|[12]\d|30|31)_(\d{4})'
'    Found Match '12_23_1995' at position 0.
'    Converted '12_23_1995' to 12-23-1995.
'
'    File name '01_09_1995.txt'
'    Pattern matched '(0?[1-9]|1[0-2])_(0?[1-9]|[12]\d|30|31)_(\d{4})'
'    Found Match '01_09_1995' at position 0.
'    Converted '01_09_1995' to 01-09-1995.
'
'    File name 'Whatever 01_09_1995'
'    Pattern matched '(0?[1-9]|1[0-2])_(0?[1-9]|[12]\d|30|31)_(\d{4})'
'    Found Match '01_09_1995' at position 9.
'    Converted '01_09_1995' to 01-09-1995.
'
'    File name 'Whatever01_09_1995'
'    Pattern matched '(0?[1-9]|1[0-2])_(0?[1-9]|[12]\d|30|31)_(\d{4})'
'    Found Match '01_09_1995' at position 8.
'    Converted '01_09_1995' to 01-09-1995.
'
'    File name 'Whatever01_09_1995text'
'    Pattern matched '(0?[1-9]|1[0-2])_(0?[1-9]|[12]\d|30|31)_(\d{4})'
'    Found Match '01_09_1995' at position 8.
'    Converted '01_09_1995' to 01-09-1995.
'
'    File name '1.09.95'
'    Pattern matched '(0?[1-9]|1[0-2])[\.](0?[1-9]|[12]\d|30|31)[\.](\d{2})'
'    Found Match '1.09.95' at position 0.
'    Converted '1.09.95' to 01-09-1995.
'
'    File name '1.9.95'
'    Pattern matched '(0?[1-9]|1[0-2])[\.](0?[1-9]|[12]\d|30|31)[\.](\d{2})'
'    Found Match '1.9.95' at position 0.
'    Converted '1.9.95' to 01-09-1995.
'
'    File name '5.09.95'
'    Pattern matched '(0?[1-9]|1[0-2])[\.](0?[1-9]|[12]\d|30|31)[\.](\d{2})'
'    Found Match '5.09.95' at position 0.
'    Converted '5.09.95' to 05-09-1995.
'
'    File name '05.12.95'
'    Pattern matched '(0?[1-9]|1[0-2])[\.](0?[1-9]|[12]\d|30|31)[\.](\d{2})'
'    Found Match '05.12.95' at position 0.
'    Converted '05.12.95' to 05-12-1995.
'
'    File name '12.09.95'
'    Pattern matched '(0?[1-9]|1[0-2])[\.](0?[1-9]|[12]\d|30|31)[\.](\d{2})'
'    Found Match '12.09.95' at position 0.
'    Converted '12.09.95' to 12-09-1995.
'
'    File name '19950109'
'    Pattern matched '\b(\d{4})(0[1-9]|1[0-2])(0[1-9]|[12]\d|30|31)\b'
'    Found Match '19950109' at position 0.
'    Converted '19950109' to 01-09-1995.
'
'    File name '01 -1995'
'    Pattern matched '(0?[1-9]|1[0-2]) -(\d{4})'
'    Found Match '01 -1995' at position 0.
'    Converted '01 -1995' to 01-01-1995.
'
'    File name 'January 1995'
'    Pattern matched '\w[a-zA-Z]+\s\d{4}'
'    Found Match 'January 1995' at position 0.
'    Converted 'January 1995' to 01-01-1995.
'
'    File name 'December 1995'
'    Pattern matched '\w[a-zA-Z]+\s\d{4}'
'    Found Match 'December 1995' at position 0.
'    Converted 'December 1995' to 12-01-1995.


