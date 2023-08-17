Attribute VB_Name = "CultureInfoCreate2Example"
'@Folder "Examples.System.Globalization.CultureInfo.Constructors"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 9, 2023
'@LastModified August 13, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.-ctor?view=netframework-4.8.1#system-globalization-cultureinfo-ctor(system-string)

Option Explicit

Public Sub CultureInfoConstructorExample()
    Dim current As DotNetLib.CultureInfo
    Set current = CultureInfo.CurrentCulture
    Debug.Print "The current culture is "; current.Name
    Dim newCulture As DotNetLib.CultureInfo
    If (current.Name = ("fr-FR")) Then
        Set newCulture = CultureInfo.Create2("fr-LU")
    Else
        Set newCulture = CultureInfo.Create2("fr-FR")
    End If
    Set CultureInfo.CurrentCulture = newCulture
    Debug.Print "The current culture is "; CultureInfo.CurrentCulture.Name
    
    ' Restore the original culture
    Set CultureInfo.CurrentCulture = current
End Sub

' The example displays output like the following:
'     The current culture is en-US
'     The current culture is now fr-FR
