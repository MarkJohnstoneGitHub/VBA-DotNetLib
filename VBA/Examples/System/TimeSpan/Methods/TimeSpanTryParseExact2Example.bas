Attribute VB_Name = "TimeSpanTryParseExact2Example"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 15, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.tryparseexact?view=netframework-4.8.1#system-timespan-tryparseexact(system-string-system-string()-system-iformatprovider-system-timespan@)
Option Explicit

''
' The following example calls the TryParseExact(String, String[], IFormatProvider, TimeSpan)
' method to convert each element of a string array to a TimeSpan value.
' The example interprets the strings by using the formatting conventions of the
' French - France ("fr-FR") culture. The strings can represent a time interval
' in either the general short format or the general long format.
''
Public Sub TimeSpanTryParseExact2()
    Dim inputs() As String
    inputs = StringArray.CreateInitialize1D("3", "16:42", "1:6:52:35.0625", _
                            "1:6:52:35,0625")
    Dim formats() As String
    formats = StringArray.CreateInitialize1D("g", "G", "%h")
    Dim interval As DotNetLib.TimeSpan
    Dim culture As DotNetLib.CultureInfo
    Set culture = CultureInfo.CreateFromName("fr-FR")
    
    ' Parse each string in inputs using formats and the fr-FR culture.
    Dim varInput As Variant
    For Each varInput In inputs
        If (TimeSpan.TryParseExact2(varInput, formats, culture, interval)) Then
            Debug.Print VBString.Format("{0} --> {1:c}", varInput, interval)
        Else
            Debug.Print VBString.Format("Unable to parse {0}", varInput)
        End If
    Next
End Sub

' The example displays the following output:
'       3 --> 03:00:00
'       16:42 --> 16:42:00
'       Unable to parse 1:6:52:35.0625
'       1:6:52:35,0625 --> 1.06:52:35.0625000


