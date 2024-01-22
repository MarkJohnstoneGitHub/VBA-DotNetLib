Attribute VB_Name = "CultureInfoCultureTypesExample"
'@Folder "Examples.System.Globalization.CultureInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 9, 2023
'@LastModified August 9, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.culturetypes?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example demonstrates the CultureTypes enumeration and the CultureTypes property.")
Public Sub CultureInfoCultureTypes()
Attribute CultureInfoCultureTypes.VB_Description = "The following example demonstrates the CultureTypes enumeration and the CultureTypes property."
    Dim cultures() As DotNetLib.CultureInfo
    cultures = CultureInfo.GetCultures(CultureTypes.CultureTypes_AllCultures)
    Dim varCultureInfo As Variant
    For Each varCultureInfo In cultures
        Dim culture As DotNetLib.CultureInfo
        Set culture = varCultureInfo
        Debug.Print culture.EnglishName; "  ("; culture.name; "):";
        If (culture.IsNeutralCulture) Then
            Debug.Print " NeutralCulture"
        Else
            Debug.Print " SpecificCulture"
        End If
    Next
End Sub

'/*
'The following is a portion of the output from this example.
'      Tajik (tg):  NeutralCulture
'      Tajik (Cyrillic) (tg-Cyrl):  NeutralCulture
'      Tajik (Cyrillic, Tajikistan) (tg-Cyrl-TJ):  SpecificCulture
'      Thai (TH):  NeutralCulture
'      Thai (Thailand) (th-TH):  SpecificCulture
'      Tigrinya (ti):  NeutralCulture
'      Tigrinya (Eritrea) (ti-ER):  SpecificCulture
'      Tigrinya (Ethiopia) (ti-ET):  SpecificCulture
'      Tigre (tig):  NeutralCulture
'      Tigre (Eritrea) (tig-ER):  SpecificCulture
'      Turkmen (tk):  NeutralCulture
'      Turkmen (Turkmenistan) (tk-TM):  SpecificCulture
'      Setswana (tn):  NeutralCulture
'      Setswana (Botswana) (tn-BW):  SpecificCulture
'      Setswana (South Africa) (tn-ZA):  SpecificCulture
'*/
