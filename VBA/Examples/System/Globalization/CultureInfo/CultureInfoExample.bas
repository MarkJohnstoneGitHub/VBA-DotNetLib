Attribute VB_Name = "CultureInfoExample"
'@Folder "Examples.System.Globalization.CultureInfo"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 9, 2023
'@LastModified September 2, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo?view=netframework-4.8.1

Option Explicit

Public Sub TestCultureInfo()
    Dim culture As ICultureInfo
    Set culture = CultureInfo.CreateFromName("fr-FR")
    Debug.Print culture.NativeName
    Debug.Print culture.NumberFormat.CurrencySymbol
    Debug.Print culture.NumberFormat.CurrencyDecimalSeparator
    
End Sub

'https://learn.microsoft.com/en-us/dotnet/api/system.globalization.culturetypes?view=netframework-4.8.1#examples
'https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.isneutralculture?view=netframework-4.8.1#examples
Public Sub TestCultureInfo2()
    Dim cultures() As DotNetLib.CultureInfo
    cultures = CultureInfo.GetCultures(CultureTypes.CultureTypes_AllCultures)
    Dim varCultureInfo As Variant
    For Each varCultureInfo In cultures
        Dim culture As DotNetLib.CultureInfo
        Set culture = varCultureInfo
        Debug.Print culture.EnglishName; "  ("; culture.Name; "):";
        If (culture.IsNeutralCulture) Then
            Debug.Print " NeutralCulture"
        Else
            Debug.Print " SpecificCulture"
        End If
    Next
End Sub

