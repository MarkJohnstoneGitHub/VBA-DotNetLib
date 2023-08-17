Attribute VB_Name = "TimeSpanTryParseExact4Example"
'@Folder("Examples.System.TimeSpan.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 16, 2023
'@LastModified August 16, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.tryparseexact?view=netframework-4.8.1#system-timespan-tryparseexact(system-string-system-string()-system-iformatprovider-system-globalization-timespanstyles-system-timespan@)

Option Explicit

' The following example calls the TryParseExact(String, String[], IFormatProvider, TimeSpanStyles, TimeSpan)
' method to convert each element of a string array to a TimeSpan value.
' The strings can represent a time interval in either the general short format or the general long format.
Public Sub TimeSpanTryParseExact4()
    Dim inputs() As String
    inputs = Strings.ToArray("3", "16:42", "1:6:52:35.0625", _
                            "1:6:52:35,0625")
    Dim formats() As String
    formats = Strings.ToArray("%h", "g", "G")
    Dim interval As DotNetLib.TimeSpan
    Dim culture As DotNetLib.CultureInfo
    Set culture = CultureInfo.Create2("fr-FR")
    
    ' Parse each string in inputs using formats and the fr-FR culture.
    Dim varInput As Variant
    For Each varInput In inputs
    
        If (TimeSpan.TryParseExact4(varInput, formats, culture, TimeSpanStyles.TimeSpanStyles_AssumeNegative, interval)) Then
            Debug.Print varInput; " --> "; interval.ToString2("c")
        Else
            Debug.Print "Unable to parse "; varInput
        End If
    Next
End Sub

' The example displays the following output:
'       3 --> -03:00:00
'       16:42 --> 16:42:00
'       Unable to parse 1:6:52:35.0625
'       1:6:52:35,0625 --> 1.06:52:35.0625000
