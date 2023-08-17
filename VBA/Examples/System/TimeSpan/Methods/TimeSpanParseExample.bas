Attribute VB_Name = "TimeSpanParseExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 17, 2023
'@LastModified August 4, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.parse?view=netframework-4.8.1#system-timespan-parse(system-string)

Option Explicit

'@Description("The following example uses the Parse method to convert each element in a string array to a TimeSpan value.")
Public Sub TimeSpanParse()
Attribute TimeSpanParse.VB_Description = "The following example uses the Parse method to convert each element in a string array to a TimeSpan value."
    Dim values() As String
    values = Strings.ToArray("6", "6:12", "6:12:14", "6:12:14:45", _
                            "6.12:14:45", "6:12:14:45.3448", _
                            "6:12:14:45,3448", "6:34:14:45")

    Dim value As Variant
    For Each value In values
        On Error Resume Next
        Dim ts As ITimeSpan
        Set ts = TimeSpan.Parse(value)
        If Try Then
            Debug.Print value & " --> " & ts.ToString2("c")
        Else
            If Catch(FormatException) Then
               Debug.Print value & ": Bad Format"
            ElseIf Catch(OverflowException) Then
               Debug.Print value & ": Overflow"
            End If
        End If
        On Error GoTo 0 'Stop code and display error
    Next
End Sub

' The example displays the following output:
'6 --> 6.00:00:00
'6:12 --> 06:12:00
'6:12:14 --> 06:12:14
'6:12:14:45 --> 6.12:14:45
'6.12:14:45 --> 6.12:14:45
'6:12:14:45.3448 --> 6.12:14:45.3448000
'6:12:14:45,3448: Bad Format
'6:34:14:45: Overflow
