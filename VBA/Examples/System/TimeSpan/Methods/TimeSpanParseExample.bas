Attribute VB_Name = "TimeSpanParseExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 17, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.parse?view=netframework-4.8.1#system-timespan-parse(system-string)

Option Explicit

''
' The following example uses the Parse method to convert each element in a
' string array to a TimeSpan value.
''
Public Sub TimeSpanParse()
    Dim values() As String
    values = StringArray.CreateInitialize1D("6", "6:12", "6:12:14", "6:12:14:45", _
                            "6.12:14:45", "6:12:14:45.3448", _
                            "6:12:14:45,3448", "6:34:14:45")
                            
    Dim cultureNames() As String
    cultureNames = StringArray.CreateInitialize1D("hr-HR", "en-US")

    ' Change the current culture.
    Dim cultureName As Variant
    For Each cultureName In cultureNames
        Set CultureInfo.CurrentCulture = CultureInfo.CreateFromName(cultureName)
        Debug.Print VBString.Format("Current Culture: {0}", _
                           CultureInfo.CurrentCulture.name)
        Dim value As Variant
        For Each value In values
            On Error Resume Next
            Dim ts As DotNetLib.TimeSpan
            Set ts = TimeSpan.Parse(value)
            If Try Then
                Debug.Print VBString.Format("{0} --> {1}", value, ts.ToString2("c"))
            Else
                If Catch(FormatException) Then
                    Debug.Print VBString.Format("{0}: Bad Format", value)
                ElseIf Catch(OverflowException) Then
                    Debug.Print VBString.Format("{0}: Overflow", value)
                End If
            End If
            On Error GoTo 0 'Stop code and display error
        Next
        Debug.Print
    Next

End Sub

' The example displays the following output:
'    Current Culture: hr-HR
'    6 --> 6.00:00:00
'    6:12 --> 06:12:00
'    6:12:14 --> 06:12:14
'    6:12:14:45 --> 6.12:14:45
'    6.12:14:45 --> 6.12:14:45
'    6:12:14:45.3448: Bad Format
'    6:12:14:45,3448 --> 6.12:14:45.3448000
'    6:34:14:45: Overflow
'
'    Current Culture: en-US
'    6 --> 6.00:00:00
'    6:12 --> 06:12:00
'    6:12:14 --> 06:12:14
'    6:12:14:45 --> 6.12:14:45
'    6.12:14:45 --> 6.12:14:45
'    6:12:14:45.3448 --> 6.12:14:45.3448000
'    6:12:14:45,3448: Bad Format
'    6:34:14:45: Overflow

