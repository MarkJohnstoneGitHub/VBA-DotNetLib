Attribute VB_Name = "CultureInfoEnglishNameExample"
'@Folder "Examples.System.Globalization.CultureInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 22, 2023
'@LastModified September 22, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.englishname?view=netframework-4.8.1#examples

Option Explicit

'The following code example displays several properties of the neutral cultures.
Public Sub CultureInfoEnglishName()
    Dim cultures() As DotNetLib.CultureInfo
    cultures = CultureInfo.GetCultures(CultureTypes.CultureTypes_NeutralCultures)
    Dim varCultureInfo As Variant
    Debug.Print "CULTURE ISO ISO WIN DISPLAYNAME                              ENGLISHNAME"
    For Each varCultureInfo In CultureInfo.GetCultures(CultureTypes.CultureTypes_NeutralCultures)
        Dim ci As DotNetLib.CultureInfo
        Set ci = varCultureInfo
        Debug.Print ci.name; " ";
        Debug.Print ci.TwoLetterISOLanguageName; " ";
        Debug.Print ci.ThreeLetterISOLanguageName; " ";
        Debug.Print ci.ThreeLetterWindowsLanguageName; " ";
        Debug.Print ci.DisplayName; " ";
        Debug.Print ci.EnglishName
    Next
End Sub

'/*
'This code produces the following output.  This output has been cropped for brevity.
'
'CULTURE ISO ISO WIN DISPLAYNAME                              ENGLISHNAME
'ar      ar  ara ARA Arabic                                   Arabic
'bg      bg  bul BGR Bulgarian                                Bulgarian
'ca      ca  cat CAT Catalan                                  Catalan
'zh-Hans zh  zho CHS Chinese (Simplified)                     Chinese (Simplified)
'cs      cs  ces CSY Czech                                    Czech
'da      da  dan DAN Danish                                   Danish
'de      de  deu DEU German                                   German
'el      el  ell ELL Greek                                    Greek
'en      en  eng ENU English                                  English
'es      es  spa ESP Spanish                                  Spanish
'fi      fi  fin FIN Finnish                                  Finnish
'zh      zh  zho CHS Chinese                                  Chinese
'zh-Hant zh  zho CHT Chinese (Traditional)                    Chinese (Traditional)
'zh-CHS  zh  zho CHS Chinese (Simplified) Legacy              Chinese (Simplified) Legacy
'zh-CHT  zh  zho CHT Chinese (Traditional) Legacy             Chinese (Traditional) Legacy
'
'*/
