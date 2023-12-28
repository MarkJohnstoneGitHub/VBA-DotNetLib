Attribute VB_Name = "StringsFormatExample"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 23, 2023
'@LastModified September 25, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.format?view=netframework-4.8.1#system-string-format(system-string-system-object())

Option Explicit

''
' This example creates a string that contains data on the high and low
' temperature on a particular date. The composite format string has five format
' items in the C# example and six in the Visual Basic example. Two of the format
' items define the width of their corresponding value's string representation,
' and the first format item also includes a standard date and time format string.
''
Public Sub StringsFormat()
    Dim date1 As DotNetLib.DateTime
    Set date1 = DateTime.CreateFromDate(2009, 7, 1)
    Dim hiTime As DotNetLib.TimeSpan
    Set hiTime = TimeSpan.Create(14, 17, 32)
    Dim hiTemp As Variant
    hiTemp = CDec(62.1)
    
    Dim loTime As DotNetLib.TimeSpan
    Set loTime = TimeSpan.Create(3, 16, 10)
    Dim loTemp As Variant
    loTemp = CDec(54.8)
    
    Dim result1 As String
    
    result1 = VBAString.Format(Regex.Unescape("Temperature on {0:d}:\n{1,11}: {2} degrees (hi)\n{3,11}: {4} degrees (lo)"), _
                           date1, hiTime, hiTemp, loTime, loTemp)
    Debug.Print result1
    Debug.Print
    
    Dim result2 As String
    result2 = VBAString.Format(Regex.Unescape("Temperature on {0:d}:\n{1,11}: {2} degrees (hi)\n{3,11}: {4} degrees (lo)"), _
                            date1, hiTime, hiTemp, loTime, loTemp)
    Debug.Print result2
End Sub

' The example displays output like the following:
'       Temperature on 7/1/2009:
'          14:17:32: 62.1 degrees (hi)
'         03:16:10: 54.8 degrees (lo)
'       Temperature on 7/1/2009:
'          14:17:32: 62.1 degrees (hi)
'          03:16:10: 54.8 degrees (lo)


