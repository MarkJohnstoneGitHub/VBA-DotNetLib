Attribute VB_Name = "TimeSpanToString3Example"
'@Folder("Examples.System.TimeSpan.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 16, 2023
'@LastModified September 2, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.tostring?view=netframework-4.8.1#system-timespan-tostring(system-string-system-iformatprovider)

Option Explicit

'@Description("The following example calls the ToString(String, IFormatProvider) method to format two time intervals.")
' The example calls the method twice for each format string, first to display
' it using the conventions of the en-US culture and then to display it using
' the conventions of the fr-FR culture.
Public Sub TimeSpanToString3()
Attribute TimeSpanToString3.VB_Description = "The following example calls the ToString(String, IFormatProvider) method to format two time intervals."
   ' Create an array of timespan intervals.
   Dim intervals() As DotNetLib.TimeSpan
   Objects.ToArray intervals, _
        TimeSpan.Create(38, 30, 15), _
        TimeSpan.Create(16, 14, 30)
    Dim cultures() As DotNetLib.CultureInfo
    Objects.ToArray cultures, _
        CultureInfo.CreateFromName("en-US"), _
        CultureInfo.CreateFromName("fr-FR")
    Dim fmts() As String
    fmts = Strings.ToArray("c", "g", "G", "hh\:mm\:ss")
    Debug.Print "Interval"; "      Format  "; cultures(0).Name; "  "; cultures(1).Name

    Dim varInterval As Variant
    For Each varInterval In intervals
        Dim interval As ITimeSpan
        Set interval = varInterval
        Dim fmt As Variant
        For Each fmt In fmts
           Debug.Print interval.ToString(); "  "; fmt; "      "; _
                        interval.ToString3(fmt, cultures(0)); "     "; _
                        interval.ToString3(fmt, cultures(1))
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
