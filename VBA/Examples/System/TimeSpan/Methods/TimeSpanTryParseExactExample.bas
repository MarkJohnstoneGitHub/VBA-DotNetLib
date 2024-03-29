Attribute VB_Name = "TimeSpanTryParseExactExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 15, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.tryparseexact?view=netframework-4.8.1#system-timespan-tryparseexact(system-string-system-string-system-iformatprovider-system-timespan@)

Option Explicit

''
' The following example uses the
' TryParseExact(String, String, IFormatProvider, TimeSpanStyles, TimeSpan) method
' to parse several string representations of time intervals using various format
' strings and cultures.
''
Public Sub TimeSpanTryParseExact()
    Dim intervalString As String
    Dim Format As String
    Dim culture As DotNetLib.CultureInfo
    
    ' Parse hour:minute value with "g" specifier current culture.
    intervalString = "17:14"
    Format = "g"
    Set culture = CultureInfo.CurrentCulture
    Dim interval As DotNetLib.TimeSpan
    If (TimeSpan.TryParseExact(intervalString, Format, culture, interval)) Then
        Debug.Print VBString.Format("'{0}' --> {1}", intervalString, interval)
    Else
        Debug.Print VBString.Format("Unable to parse {0}", intervalString)
    End If

    ' Parse hour:minute:second value with "G" specifier.
    intervalString = "17:14:48"
    Format = "G"
    Set culture = CultureInfo.InvariantCulture
    If (TimeSpan.TryParseExact(intervalString, Format, culture, interval)) Then
        Debug.Print VBString.Format("'{0}' --> {1}", intervalString, interval)
    Else
        Debug.Print VBString.Format("Unable to parse {0}", intervalString)
    End If
    
    ' Parse hour:minute:second value with "G" specifier.
    ' and current (en-US) culture.
    intervalString = "17:14:48.153"
    Format = "G"
    Set culture = CultureInfo.InvariantCulture
    If (TimeSpan.TryParseExact(intervalString, Format, culture, interval)) Then
        Debug.Print VBString.Format("'{0}' --> {1}", intervalString, interval)
    Else
        Debug.Print VBString.Format("Unable to parse {0}", intervalString)
    End If

    ' Parse days:hours:minute.second value with "G" specifier
    ' and current (en-US) culture.
    intervalString = "3:17:14:48.153"
    Format = "G"
    Set culture = CultureInfo.CurrentCulture
    If (TimeSpan.TryParseExact(intervalString, Format, culture, interval)) Then
        Debug.Print VBString.Format("'{0}' --> {1}", intervalString, interval)
    Else
        Debug.Print VBString.Format("Unable to parse {0}", intervalString)
    End If
    
    ' Parse days:hours:minute.second value with "G" specifier
    ' and fr-FR culture.
    intervalString = "3:17:14:48.153"
    Format = "G"
    Set culture = CultureInfo.CreateFromName("fr-FR")
    If (TimeSpan.TryParseExact(intervalString, Format, culture, interval)) Then
        Debug.Print VBString.Format("'{0}' --> {1}", intervalString, interval)
    Else
        Debug.Print VBString.Format("Unable to parse {0}", intervalString)
    End If

    ' Parse a single number using the "c" standard format string.
    intervalString = "12"
    Format = "c"
    If (TimeSpan.TryParseExact(intervalString, Format, Nothing, interval)) Then
        Debug.Print VBString.Format("'{0}' --> {1}", intervalString, interval)
    Else
        Debug.Print VBString.Format("Unable to parse {0}", intervalString)
    End If

    ' Parse a single number using the "%h" custom format string.
    Format = "%h"
    If (TimeSpan.TryParseExact(intervalString, Format, Nothing, interval)) Then
        Debug.Print VBString.Format("'{0}' --> {1}", intervalString, interval)
    Else
        Debug.Print VBString.Format("Unable to parse {0}", intervalString)
    End If

    ' Parse a single number using the "%s" custom format string.
    Format = "%s"
    If (TimeSpan.TryParseExact(intervalString, Format, Nothing, interval)) Then
        Debug.Print VBString.Format("'{0}' --> {1}", intervalString, interval)
    Else
        Debug.Print VBString.Format("Unable to parse {0}", intervalString)
    End If

End Sub

' The example displays the following output:
'       '17:14' --> 17:14:00
'       Unable to parse 17:14:48
'       Unable to parse 17:14:48.153
'       '3:17:14:48.153' --> 3.17:14:48.1530000
'       Unable to parse 3:17:14:48.153
'       '3:17:14:48,153' --> 3.17:14:48.1530000
'       '12' --> 12.00:00:00
'       '12' --> 12:00:00
'       '12' --> 00:00:12
