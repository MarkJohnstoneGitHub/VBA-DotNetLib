Attribute VB_Name = "VBStringFormat5Example"
'@Folder "Examples.System.VBString.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 23, 2023
'@LastModified January 23, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.format?view=netframework-4.8.1#system-string-format(system-iformatprovider-system-string-system-object())

Option Explicit

''
' This example uses the Format(IFormatProvider, String, Object[]) method to
' display the string representation of some date and time values and numeric
' values by using several different cultures.
''
Public Sub VBStringFormat5Example()
    Dim cultureNames() As String
    cultureNames = StringArray.CreateInitialize1D("en-US", "fr-FR", "de-DE", "es-ES")
    Dim dateToDisplay As DotNetLib.DateTime
    Set dateToDisplay = DateTime.CreateFromDateTime(2009, 9, 1, 18, 32, 0)
    Dim value As Double
    value = 9164.32
    
    Debug.Print "Culture     Date                                Value"
    Dim cultureName As Variant
    For Each cultureName In cultureNames
        Dim culture As DotNetLib.CultureInfo
        Set culture = CultureInfo.CreateFromName(cultureName)
        Dim Output As String
        Output = VBString.Format5(culture, "{0,-11} {1,-35:D} {2:N}", _
                                culture.name, dateToDisplay, value)
        Debug.Print Output
    Next
End Sub

' The example displays the following output:
'    Culture     Date                                Value
'
'    en-US       Tuesday, September 01, 2009         9,164.32
'    fr-FR       mardi 1 septembre 2009              9 164,32
'    de-DE       Dienstag, 1. September 2009         9.164,32
'    es-ES       martes, 01 de septiembre de 2009    9.164,32


