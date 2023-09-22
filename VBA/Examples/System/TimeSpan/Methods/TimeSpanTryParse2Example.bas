Attribute VB_Name = "TimeSpanTryParse2Example"
'@Folder("Examples.System.TimeSpan.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 15, 2023
'@LastModified September 2, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.tryparse?view=netframework-4.8.1#system-timespan-tryparse(system-string-system-iformatprovider-system-timespan@)

Option Explicit

' The following example defines an array of CultureInfo objects, and uses each
' object in calls to the TryParse(String, IFormatProvider, TimeSpan) method to
' parse the elements in a string array. The example illustrates how the conventions
' of a specific culture influence the formatting operation.
Public Sub TimeSpanTryParse2()
    ' Create an array of all supported standard date and time format specifiers.
    Dim values() As String
    values = StringArray.ToArray("6", "6:12", "6:12:14", "6:12:14:45", _
                            "6.12:14:45", "6:12:14:45.3448", _
                            "6:12:14:45,3448", "6:34:14:45")
    
    ' Create an array of four cultures.
    Dim cultures() As DotNetLib.CultureInfo
    Objects.ToArray cultures, _
                    CultureInfo.CreateFromName("en-US"), _
                    CultureInfo.CreateFromName("ru-RU"), _
                    CultureInfo.InvariantCulture
    
    Dim header As String
    header = "String           "
    Dim varCulture As Variant
    Dim culture As DotNetLib.CultureInfo
    For Each varCulture In cultures
        Set culture = varCulture
        header = header & "     " & IIf(culture.Equals(CultureInfo.InvariantCulture), "Invariant", culture.Name)
    Next
    Debug.Print header
    Debug.Print
    
    Dim value As Variant
    For Each value In values
        Debug.Print value; "           ";
        For Each varCulture In cultures
            Set culture = varCulture
            Dim interval As DotNetLib.TimeSpan
            If (TimeSpan.TryParse2(value, culture, interval)) Then
                Debug.Print interval.ToString2("c"); "         ";
            Else
                Debug.Print "Unable to Parse"; "         ";
            End If
        Next
        Debug.Print
        
    Next
End Sub

' The example displays the following output:
'    String                          en-US               ru-RU           Invariant
'
'    6                          6.00:00:00          6.00:00:00          6.00:00:00
'    6:12                         06:12:00            06:12:00            06:12:00
'    6:12:14                      06:12:14            06:12:14            06:12:14
'    6:12:14:45                 6.12:14:45          6.12:14:45          6.12:14:45
'    6.12:14:45                 6.12:14:45          6.12:14:45          6.12:14:45
'    6:12:14:45.3448    6.12:14:45.3448000     Unable to Parse  6.12:14:45.3448000
'    6:12:14:45,3448       Unable to Parse  6.12:14:45.3448000     Unable to Parse
'    6:34:14:45            Unable to Parse     Unable to Parse     Unable to Parse


