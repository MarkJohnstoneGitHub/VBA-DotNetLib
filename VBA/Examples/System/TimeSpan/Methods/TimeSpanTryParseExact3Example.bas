Attribute VB_Name = "TimeSpanTryParseExact3Example"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 16, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.tryparseexact?view=netframework-4.8.1#system-timespan-tryparseexact(system-string-system-string-system-iformatprovider-system-globalization-timespanstyles-system-timespan@)

Option Explicit

''
' The following example uses the ParseExact(String, String, IFormatProvider)
' method to parse several string representations of time intervals using various
' format strings and cultures. It also uses the TimeSpanStyles.AssumeNegative
' value to interpret each string as a negative time interval.
' The output from the example illustrates that the TimeSpanStyles.AssumeNegative
' style affects the return value only when it is used with custom format strings.
''
Public Sub TimeSpanTryParseExact3()
    Dim intervalString As String
    Dim pvtFormat As String
    Dim interval As DotNetLib.TimeSpan
    Dim culture As DotNetLib.CultureInfo
    Set culture = Nothing
    
    ' Parse hour:minute value with custom format specifier.
    intervalString = "17:14"
    pvtFormat = "hh\:mm"
    Set culture = CultureInfo.CurrentCulture
    If (TimeSpan.TryParseExact3(intervalString, pvtFormat, culture, TimeSpanStyles.TimeSpanStyles_AssumeNegative, interval)) Then
        Debug.Print VBString.Format("'{0}' ({1}) --> {2}", intervalString, pvtFormat, interval)
    Else
        Debug.Print VBString.Format("Unable to parse '{0}' using format {1}", _
                        intervalString, pvtFormat)
    End If

    ' Parse hour:minute:second value with "g" specifier.
    intervalString = "17:14:48"
    pvtFormat = "g"
    Set culture = CultureInfo.InvariantCulture
    If (TimeSpan.TryParseExact3(intervalString, pvtFormat, culture, TimeSpanStyles.TimeSpanStyles_AssumeNegative, interval)) Then
        Debug.Print VBString.Format("'{0}' ({1}) --> {2}", intervalString, pvtFormat, interval)
    Else
        Debug.Print VBString.Format("Unable to parse '{0}' using format {1}", _
                        intervalString, pvtFormat)
    End If

    ' Parse hours:minute.second value with custom format specifier.
    intervalString = "17:14:48.153"
    pvtFormat = "h\:mm\:ss\.fff"
    Set culture = Nothing
    If (TimeSpan.TryParseExact3(intervalString, pvtFormat, culture, TimeSpanStyles.TimeSpanStyles_AssumeNegative, interval)) Then
        Debug.Print VBString.Format("'{0}' ({1}) --> {2}", intervalString, pvtFormat, interval)
    Else
        Debug.Print VBString.Format("Unable to parse '{0}' using format {1}", _
                        intervalString, pvtFormat)
    End If

    ' Parse days:hours:minute.second value with "G" specifier
    ' and current (en-US) culture.
    intervalString = "3:17:14:48.153"
    pvtFormat = "G"
    Set culture = CultureInfo.CurrentCulture
    If (TimeSpan.TryParseExact3(intervalString, pvtFormat, culture, TimeSpanStyles.TimeSpanStyles_AssumeNegative, interval)) Then
        Debug.Print VBString.Format("'{0}' ({1}) --> {2}", intervalString, pvtFormat, interval)
    Else
        Debug.Print VBString.Format("Unable to parse '{0}' using format {1}", _
                        intervalString, pvtFormat)
    End If

    ' Parse days:hours:minute.second value with a custom format specifier.
    intervalString = "3:17:14:48.153"
    pvtFormat = "d\:hh\:mm\:ss\.fff"
    Set culture = Nothing
    If (TimeSpan.TryParseExact3(intervalString, pvtFormat, culture, TimeSpanStyles.TimeSpanStyles_AssumeNegative, interval)) Then
        Debug.Print VBString.Format("'{0}' ({1}) --> {2}", intervalString, pvtFormat, interval)
    Else
        Debug.Print VBString.Format("Unable to parse '{0}' using format {1}", _
                        intervalString, pvtFormat)
    End If

    ' Parse days:hours:minute.second value with "G" specifier
    ' and fr-FR culture.
    intervalString = "3:17:14:48,153"
    pvtFormat = "G"
    Set culture = CultureInfo.CreateFromName("fr-FR")
    If (TimeSpan.TryParseExact3(intervalString, pvtFormat, culture, TimeSpanStyles.TimeSpanStyles_AssumeNegative, interval)) Then
        Debug.Print VBString.Format("'{0}' ({1}) --> {2}", intervalString, pvtFormat, interval)
    Else
        Debug.Print VBString.Format("Unable to parse '{0}' using format {1}", _
                        intervalString, pvtFormat)
    End If

    ' Parse a single number using the "c" standard format string.
    intervalString = "12"
    pvtFormat = "c"
    If (TimeSpan.TryParseExact3(intervalString, pvtFormat, Nothing, TimeSpanStyles.TimeSpanStyles_AssumeNegative, interval)) Then
        Debug.Print VBString.Format("'{0}' ({1}) --> {2}", intervalString, pvtFormat, interval)
    Else
        Debug.Print VBString.Format("Unable to parse '{0}' using format {1}", _
                        intervalString, pvtFormat)
    End If
    
    ' Parse a single number using the "%h" custom format string.
    pvtFormat = "%h"
    If (TimeSpan.TryParseExact3(intervalString, pvtFormat, Nothing, TimeSpanStyles.TimeSpanStyles_AssumeNegative, interval)) Then
        Debug.Print VBString.Format("'{0}' ({1}) --> {2}", intervalString, pvtFormat, interval)
    Else
        Debug.Print VBString.Format("Unable to parse '{0}' using format {1}", _
                        intervalString, pvtFormat)
    End If
    
    ' Parse a single number using the "%s" custom format string.
    pvtFormat = "%s"
    If (TimeSpan.TryParseExact3(intervalString, pvtFormat, Nothing, TimeSpanStyles.TimeSpanStyles_AssumeNegative, interval)) Then
        Debug.Print VBString.Format("'{0}' ({1}) --> {2}", intervalString, pvtFormat, interval)
    Else
        Debug.Print VBString.Format("Unable to parse '{0}' using format {1}", _
                        intervalString, pvtFormat)
    End If
End Sub

' The example displays the following output:
'    '17:14' (h\:mm) --> -17:14:00
'    '17:14:48' (g) --> 17:14:48
'    '17:14:48.153' (h\:mm\:ss\.fff) --> -17:14:48.1530000
'    '3:17:14:48.153' (G) --> 3.17:14:48.1530000
'    '3:17:14:48.153' (d\:hh\:mm\:ss\.fff) --> -3.17:14:48.1530000
'    '3:17:14:48,153' (G) --> 3.17:14:48.1530000
'    '12' (c) --> 12.00:00:00
'    '12' (%h) --> -12:00:00
'    '12' (%s) --> -00:00:12


