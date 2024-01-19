Attribute VB_Name = "TimeSpanToString2Example2"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 16, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.tostring?view=netframework-4.8.1#system-timespan-tostring(system-string-system-iformatprovider)

Option Explicit

''
' The following example calls the ToString(String, IFormatProvider) method to
' format two time intervals. The example calls the method twice for each
' format string, first to display it using the conventions of the en-US
' culture and then to display it using the conventions of the fr-FR culture.
''
Public Sub TimeSpanToString2Example2()
   ' Create an array of timespan intervals.
   Dim intervals() As DotNetLib.TimeSpan
   ObjectArray.CreateInitialize1D intervals, _
        TimeSpan.Create(38, 30, 15), _
        TimeSpan.Create(16, 14, 30)
    Dim cultures() As DotNetLib.CultureInfo
    ObjectArray.CreateInitialize1D cultures, _
        CultureInfo.CreateFromName("en-US"), _
        CultureInfo.CreateFromName("fr-FR")
    Dim fmts() As String
    fmts = StringArray.CreateInitialize1D("c", "g", "G", "hh\:mm\:ss")
    Debug.Print VBString.Format(VBString.Unescape("{0,12}      Format  {1,22}  {2,22}\n"), _
                      "Interval", cultures(0).Name, cultures(1).Name)

    Dim varInterval As Variant
    For Each varInterval In intervals
        Dim interval As DotNetLib.TimeSpan
        Set interval = varInterval
        Dim fmt As Variant
        For Each fmt In fmts
            Debug.Print VBString.Format("{0,12}  {1,10}  {2,22}  {3,22}", _
                            interval, fmt, _
                            interval.ToString2(fmt, cultures(0)), _
                            interval.ToString2(fmt, cultures(1)))
        Next
        Debug.Print
    Next
End Sub

' The example displays the following output:
'        Interval      Format                   en-US                   fr-FR
'
'      1.14:30:15           c              1.14:30:15              1.14:30:15
'      1.14:30:15           g              1:14:30:15              1:14:30:15
'      1.14:30:15           G      1:14:30:15.0000000      1:14:30:15,0000000
'      1.14:30:15  hh\:mm\:ss                14:30:15                14:30:15
'
'        16:14:30           c                16:14:30                16:14:30
'        16:14:30           g                16:14:30                16:14:30
'        16:14:30           G      0:16:14:30.0000000      0:16:14:30,0000000
'        16:14:30  hh\:mm\:ss                16:14:30                16:14:30

