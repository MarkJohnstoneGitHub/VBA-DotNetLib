Attribute VB_Name = "DateTimeParse2Example"
'@Folder("Examples.System.DateTime.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 18, 2023
'@LastModified September 9, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.parse?view=netframework-4.8.1#system-datetime-parse(system-string-system-iformatprovider)

Option Explicit

' The following example parses an array of date strings by using the conventions
' of the en-US, fr-FR, and de-DE cultures. It demonstrates that the string
' representations of a single date can be interpreted differently across
' different cultures.
Public Sub DateTimeParse2()
    ' Define cultures to be used to parse dates.
    Dim cultures() As DotNetLib.CultureInfo
    Objects.ToArray cultures, _
                    CultureInfo.CreateSpecificCulture("en-US"), _
                    CultureInfo.CreateSpecificCulture("fr-FR"), _
                    CultureInfo.CreateSpecificCulture("de-DE")
    ' Define string representations of a date to be parsed.
    Dim dateStrings() As String
    dateStrings = StringArray.ToArray( _
                    "01/10/2009 7:34 PM", _
                    "10.01.2009 19:34", _
                    "10-1-2009 19:34")
    
    ' Parse dates using each culture.
    Dim varCulture As Variant
    Dim culture As DotNetLib.CultureInfo
    For Each varCulture In cultures
        Set culture = varCulture
        Dim dateValue As DotNetLib.DateTime
        Debug.Print "Attempted conversions using "; culture.Name; " culture."

        Dim dateString As Variant
        For Each dateString In dateStrings
            On Error Resume Next
            Set dateValue = DateTime.Parse2(dateString, culture)
            If Try Then
                Debug.Print "   Converted '"; dateString; "' to "; dateValue.ToString2("f", culture); "."
            ElseIf Catch(FormatException) Then
                Debug.Print "   Unable to convert '"; dateString; "' for culture "; culture.Name
            End If
            On Error GoTo 0 'reset error handling
        Next
        Debug.Print
    Next
End Sub

' The example displays the following output to the console:
'       Attempted conversions using en-US culture.
'          Converted '01/10/2009 7:34 PM' to Saturday, January 10, 2009 7:34 PM.
'          Converted '10.01.2009 19:34' to Thursday, October 01, 2009 7:34 PM.
'          Converted '10-1-2009 19:34' to Thursday, October 01, 2009 7:34 PM.
'
'       Attempted conversions using fr-FR culture.
'          Converted '01/10/2009 7:34 PM' to jeudi 1 octobre 2009 19:34.
'          Converted '10.01.2009 19:34' to samedi 10 janvier 2009 19:34.
'          Converted '10-1-2009 19:34' to samedi 10 janvier 2009 19:34.
'
'       Attempted conversions using de-DE culture.
'          Converted '01/10/2009 7:34 PM' to Donnerstag, 1. Oktober 2009 19:34.
'          Converted '10.01.2009 19:34' to Samstag, 10. Januar 2009 19:34.
'          Converted '10-1-2009 19:34' to Samstag, 10. Januar 2009 19:34.

